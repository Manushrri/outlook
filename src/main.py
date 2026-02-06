"""
Microsoft Outlook MCP Server - Main Entry Point
"""

import importlib
import inspect
import json
import logging
import sys
from pathlib import Path
from typing import Any, Dict, Optional

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from fastmcp import FastMCP
from fastmcp.tools.tool import FunctionTool
from dotenv import load_dotenv

# Force load .env from CWD
load_dotenv(verbose=True)

from src.config import settings
from src.client import OutlookClient

# Configure Logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s [%(name)s] %(message)s")
logger = logging.getLogger("outlook-mcp")

# Root path for manifest loading
ROOT_DIR = Path(__file__).resolve().parents[1]

# Server State
state: Dict[str, Any] = {}

mcp = FastMCP("outlook-mcp")


def get_client() -> OutlookClient:
    """Get or create the singleton OutlookClient."""
    if "client" not in state:
        logger.info("Initializing OutlookClient...")
        try:
            state["client"] = OutlookClient()
            
            # Check authentication
            if not state["client"].is_authenticated():
                logger.info("Authentication required. Starting interactive auth flow...")
                state["client"].authenticate_interactive()
                
        except Exception as e:
            logger.error(f"Failed to initialize client: {e}")
            raise
    return state["client"]


def remove_null_from_schema(schema):
    """Remove null types from schema to prevent MCP Inspector trim() errors."""
    if isinstance(schema, dict):
        new_schema = {}
        for key, value in schema.items():
            if key == "anyOf" and isinstance(value, list):
                # Remove null types from anyOf
                filtered = [v for v in value if not (isinstance(v, dict) and v.get("type") == "null")]
                if filtered:
                    if len(filtered) == 1:
                        # If only one type left, use it directly
                        new_schema.update(filtered[0])
                    else:
                        new_schema[key] = filtered
            elif key == "type" and isinstance(value, list) and "null" in value:
                # Handle type: ["string", "null"] -> type: "string"
                filtered = [v for v in value if v != "null"]
                if len(filtered) == 1:
                    new_schema["type"] = filtered[0]
                else:
                    new_schema["type"] = filtered
            elif key == "type" and value == "null":
                # Skip null types
                continue
            elif key == "default" and value is None:
                # Remove None defaults
                continue
            else:
                new_schema[key] = remove_null_from_schema(value) if isinstance(value, (dict, list)) else value
        return new_schema
    elif isinstance(schema, list):
        return [remove_null_from_schema(item) for item in schema]
    else:
        return schema


def register_tools():
    """Register tools from tools_manifest.json dynamically."""
    manifest_path = ROOT_DIR / "tools_manifest.json"
    if not manifest_path.exists():
        logger.error(f"Manifest not found at {manifest_path}")
        return

    try:
        with open(manifest_path, 'r') as f:
            manifest = json.load(f)
    except Exception as e:
        logger.error(f"Failed to load manifest: {e}")
        return

    logger.info(f"Loading tools from manifest...")
    
    tools_registered = 0
    for entry in manifest.get("tools", []):
        tool_id = entry.get("id")
        target = entry.get("target")
        description = entry.get("description")
        input_schema = entry.get("input_schema")
        
        if not tool_id or not target:
            continue

        try:
            module_name, func_name = target.split(":")
            module = importlib.import_module(module_name)
            func = getattr(module, func_name)
        except Exception as e:
            logger.error(f"Failed to import {target}: {e}")
            continue

        # Create Dynamic Wrapper
        try:
            wrapper = create_dynamic_wrapper(func, description, tool_id)
            
            # Use FunctionTool to create the tool object explicitly
            tool = FunctionTool.from_function(
                wrapper,
                name=tool_id,
                description=description
            )
            
            # Enforce sanitized schema from manifest (removing null types)
            if input_schema:
                cleaned_schema = remove_null_from_schema(input_schema)
                # Override the auto-generated schema with our clean manifest schema
                tool.parameters = cleaned_schema
            
            # Register with FastMCP
            mcp.add_tool(tool)
            tools_registered += 1
            logger.info(f"Registered tool: {tool_id}")
            
        except Exception as e:
            logger.error(f"Failed to wrap/register {tool_id}: {e}")

    logger.info(f"Total tools registered: {tools_registered}")


def create_dynamic_wrapper(func, description, tool_id=None):
    """
    Creates a wrapper function that matches the signature of `func` (minus 'client')
    and injects the client instance.
    """
    sig = inspect.signature(func)
    params = [p for p in sig.parameters.values() if p.name != "client"]
    
    # Build param string (e.g., "arg1, arg2=5")
    decl_parts = []
    names = []
    for p in params:
        if p.default is inspect._empty:
            decl_parts.append(p.name)
        else:
            decl_parts.append(f"{p.name}={repr(p.default)}")
        names.append(p.name)
    
    decl = ", ".join(decl_parts)
    
    # Source code for wrapper
    # We use __get_client() to fetch the client at runtime
    # Convert empty strings and empty dicts to None for optional parameters
    normalize_kwargs = []
    for n in names:
        # Check if parameter has a default (is optional)
        param = next((p for p in params if p.name == n), None)
        if param and param.default is not inspect._empty:
            # Optional parameter - normalize empty strings and empty dicts to None
            # Check if it's an empty string or an empty dict
            normalize_kwargs.append(f"{n!r}: None if ({n} == '' or (isinstance({n}, dict) and len({n}) == 0)) else {n}")
        else:
            # Required parameter - use as-is
            normalize_kwargs.append(f"{n!r}: {n}")
    
    src = (
        f"def wrapper({decl}):\n"
        f"    kwargs = {{{', '.join(normalize_kwargs)}}}\n"
        f"    client = __get_client()\n"
        f"    return __func(client=client, **kwargs)\n"
    )
    
    local_ns = {}
    global_ns = {
        "__func": func,
        "__get_client": get_client
    }
    
    exec(src, global_ns, local_ns)
    wrapper = local_ns["wrapper"]
    
    # Update Metadata
    wrapper.__name__ = tool_id if tool_id else func.__name__
    wrapper.__doc__ = description or func.__doc__
    
    # Update Annotations (remove client)
    if hasattr(func, "__annotations__"):
        ann = dict(func.__annotations__)
        ann.pop("client", None)
        wrapper.__annotations__ = ann
        
    return wrapper


def main():
    """Main entry point."""
    # Load tools from manifest
    register_tools()
    
    # Run server
    mcp.run()


if __name__ == "__main__":
    main()

