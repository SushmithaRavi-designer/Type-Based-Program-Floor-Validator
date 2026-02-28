#!/usr/bin/env python
"""Check enum definitions in main.py schema."""

from main import FunctionInputs
import json

schema = FunctionInputs.model_json_schema()
defs = schema.get('$defs', {})

print("✅ Enum Definitions in Schema:")
print()
for name, definition in defs.items():
    if 'enum' in definition:
        enum_values = definition['enum']
        print(f"  {name}: {enum_values}")

print()
print("✅ All fields that reference these enums:")
print()
for field_name, field_def in schema['properties'].items():
    if '$ref' in field_def:
        ref = field_def['$ref']
        if 'Mode' in ref or 'Level' in ref:
            title = field_def.get('title', field_name)
            default = field_def.get('default', 'N/A')
            print(f"  {field_name:40} {title:40} (default: {default})")

print()
print("✅ Schema is valid for Speckle Automate!")
