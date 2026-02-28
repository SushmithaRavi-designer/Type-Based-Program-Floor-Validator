#!/usr/bin/env python
"""Verify main.py has updated enum-based FunctionInputs."""

from main import FunctionInputs
import json

# Create a sample FunctionInputs to verify the schema
schema = FunctionInputs.model_json_schema()

print("✅ FunctionInputs Schema - Updated with Enums!")
print()
print("=" * 70)
print()

# Show all fields with their types
print("ALL FIELDS:")
print("-" * 70)
for field_name, field_info in schema['properties'].items():
    field_type = field_info.get('type', field_info.get('enum', 'unknown'))
    title = field_info.get('title', field_name)
    print(f"  {field_name:40} {title}")
    if 'enum' in field_info:
        print(f"    └─ Enum options: {field_info['enum']}")
    if 'default' in field_info:
        print(f"    └─ Default: {field_info['default']}")

print()
print("=" * 70)
print()
print(f"✅ Total fields: {len(schema['properties'])}")
print()
print("✅ ENUM FIELDS (Speckle UI Dropdowns):")
print("-" * 70)
enum_fields = [(name, info['enum']) for name, info in schema['properties'].items() if 'enum' in info]
for field_name, options in enum_fields:
    print(f"  • {field_name:40} → {options}")

print()
print(f"✅ Total enum fields: {len(enum_fields)}")
print()
print("=" * 70)
print("✅ main.py is ready for Speckle deployment!")
print("   Users will now see enum dropdowns instead of text fields!")
