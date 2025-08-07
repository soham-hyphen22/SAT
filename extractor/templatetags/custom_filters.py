# extractor/templatetags/custom_filters.py
import json
from django import template
from datetime import datetime, timedelta

register = template.Library()

@register.filter
def get_item(dictionary, key):
    """Get item from dictionary by key"""
    if isinstance(dictionary, dict):
        return dictionary.get(key, '')
    return ''

class CustomJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, timedelta)):
            return str(obj)
        return super().default(obj)

@register.filter
def pretty_json(value):
    """Convert a Python object to pretty-printed JSON"""
    try:
        return json.dumps(value, indent=2, ensure_ascii=False, cls=CustomJSONEncoder)
    except (TypeError, ValueError):
        return str(value)

@register.filter
def to_json(value):
    """Convert a Python object to JSON string"""
    try:
        return json.dumps(value, ensure_ascii=False, cls=CustomJSONEncoder)
    except (TypeError, ValueError):
        return str(value)