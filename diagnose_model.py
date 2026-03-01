#!/usr/bin/env python3
"""
Diagnostic script to inspect Speckle model structure.
Run this to understand what parameters and properties are in your model.
"""

import json
import inspect
from flatten import flatten_base
from specklepy.objects.base import Base


def diagnose_speckle_objects():
    """
    Shows detailed structure of available properties and parameters.
    """
    print("=" * 80)
    print("SPECKLE MODEL DIAGNOSTIC TOOL")
    print("=" * 80)
    print()
    
    # Example: If you have a JSON file or can load the model
    print("To use this script, you need to provide Speckle model data.")
    print()
    print("Example usage in your code:")
    print("""
    from diagnose_model import inspect_base_object
    
    # After receiving version_root_object
    obj = generic_models[0]  # or any Base object
    inspect_base_object(obj)
    """)
    print()


def inspect_base_object(obj, max_depth=3, depth=0):
    """
    Recursively inspect a Speckle Base object, showing all properties and parameters.
    
    Args:
        obj: Base object to inspect
        max_depth: Maximum depth to traverse
        depth: Current depth (for indentation)
    """
    indent = "  " * depth
    
    if depth == 0:
        print("=" * 80)
        print("INSPECTING SPECKLE BASE OBJECT")
        print("=" * 80)
        print()
    
    # Basic info
    obj_type = type(obj).__name__
    obj_class = f"{obj.__class__.__module__}.{obj.__class__.__name__}"
    print(f"{indent}Object Type: {obj_type}")
    print(f"{indent}Full Class: {obj_class}")
    print()
    
    # Get all attributes
    print(f"{indent}─ DIRECT ATTRIBUTES:")
    attrs = [attr for attr in dir(obj) if not attr.startswith("_")]
    for attr in attrs[:30]:  # Show first 30
        try:
            val = getattr(obj, attr, None)
            if not callable(val) and not isinstance(val, (list, dict, Base)):
                val_type = type(val).__name__
                print(f"{indent}  {attr}: {val} ({val_type})")
        except Exception as e:
            pass
    
    if len(attrs) > 30:
        print(f"{indent}  ... and {len(attrs) - 30} more attributes")
    print()
    
    # Check parameters
    params = getattr(obj, "parameters", None)
    if params:
        print(f"{indent}─ PARAMETERS OBJECT:")
        print(f"{indent}  Type: {type(params).__name__}")
        
        if isinstance(params, dict):
            print(f"{indent}  Dict keys: {list(params.keys())[:20]}")
            # Show first parameter structure
            if params:
                first_key = list(params.keys())[0]
                first_val = params[first_key]
                print(f"{indent}  Example '{first_key}':")
                print(f"{indent}    Value: {first_val}")
                print(f"{indent}    Type: {type(first_val).__name__}")
                if isinstance(first_val, dict):
                    print(f"{indent}    Keys: {list(first_val.keys())}")
        else:
            print(f"{indent}  (DynamicBase object)")
            param_attrs = [k for k in dir(params) if not k.startswith("_")]
            print(f"{indent}  Attributes: {param_attrs[:15]}")
    else:
        print(f"{indent}─ No 'parameters' object found")
    
    print()
    
    # Show nested objects
    if depth < max_depth:
        nested = []
        for attr in dir(obj):
            if not attr.startswith("_"):
                try:
                    val = getattr(obj, attr, None)
                    if isinstance(val, Base):
                        nested.append((attr, val))
                    elif isinstance(val, list):
                        for item in val[:3]:
                            if isinstance(item, Base):
                                nested.append((f"{attr}[{val.index(item)}]", item))
                except Exception:
                    pass
        
        if nested:
            print(f"{indent}─ NESTED OBJECTS (max_depth={max_depth}, current={depth}):")
            for nested_attr, nested_obj in nested[:5]:
                print(f"{indent}  -> {nested_attr}")
                inspect_base_object(nested_obj, max_depth, depth + 1)
    
    if depth == 0:
        print("=" * 80)
        print("END OF DIAGNOSTIC")
        print("=" * 80)


def show_property_sources(obj):
    """
    Show all possible ways a property could be stored in this object.
    """
    print("=" * 80)
    print("PROPERTY SOURCE ANALYSIS")
    print("=" * 80)
    print()
    
    print("Common property names and their locations:")
    print()
    
    search_names = [
        "Level", "level", "Floor", "floor", "Story", "story",
        "Zone", "zone", "Program", "program", "Type", "type", "Type Name",
        "Area", "area", "Material", "material", "Color", "color",
        "Category", "category", "Family", "family"
    ]
    
    for prop_name in search_names:
        found = False
        
        # Check direct attribute
        if hasattr(obj, prop_name):
            val = getattr(obj, prop_name, None)
            if val and not callable(val):
                print(f"FOUND: Direct attribute: {prop_name} = {val}")
                found = True
        
        # Check parameters dict
        params = getattr(obj, "parameters", None)
        if isinstance(params, dict):
            if prop_name in params:
                val = params[prop_name]
                print(f"FOUND: Parameter dict: {prop_name} = {val}")
                found = True
            # Check nested structure
            for key, entry in list(params.items())[:5]:
                if isinstance(entry, dict) and entry.get("name") == prop_name:
                    val = entry.get("value")
                    print(f"FOUND: Parameter nested: {prop_name} = {val}")
                    found = True
                    break
        
        if not found:
            print(f"NOT FOUND: {prop_name}")
    
    print()
    print("=" * 80)


if __name__ == "__main__":
    diagnose_speckle_objects()
    print()
    print("To inspect your actual model, add this to main.py:")
    print("""
    from diagnose_model import inspect_base_object, show_property_sources
    
    if generic_models:
        first_obj = generic_models[0]
        inspect_base_object(first_obj)
        show_property_sources(first_obj)
    """)
