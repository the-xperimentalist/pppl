from django import template
from decimal import Decimal

register = template.Library()

@register.filter
def smart_decimal(value, max_decimals=8):
    """
    Display decimal values intelligently:
    - Shows up to 2 decimal places for simple values
    - Shows up to 8 decimal places if value has higher precision
    - Strips trailing zeros
    
    Usage: {{ value|smart_decimal }}
    """
    if value is None or value == '':
        return '0'
    
    try:
        # Convert to Decimal for precise handling
        dec_value = Decimal(str(value))
        
        # Check if value is zero
        if dec_value == 0:
            return '0'
        
        # Convert to string and remove trailing zeros
        str_value = str(dec_value)
        
        # Split into integer and decimal parts
        if '.' in str_value:
            integer_part, decimal_part = str_value.split('.')
            
            # Remove trailing zeros from decimal part
            decimal_part = decimal_part.rstrip('0')
            
            # If no decimal part remains, return just integer
            if not decimal_part:
                return integer_part
            
            # Limit to max_decimals places
            decimal_part = decimal_part[:max_decimals]
            
            # If decimal part is 2 or fewer significant digits, show only those
            # Otherwise show all significant digits up to max_decimals
            if len(decimal_part) <= 2:
                return f"{integer_part}.{decimal_part}"
            else:
                return f"{integer_part}.{decimal_part}"
        else:
            return str_value
            
    except (ValueError, TypeError, decimal.InvalidOperation):
        return str(value)


@register.filter
def smart_decimal_fixed(value, decimals=2):
    """
    Display decimal values with a fixed number of decimal places,
    but only if the value has significant digits at that precision.
    
    Usage: {{ value|smart_decimal_fixed:4 }}
    """
    if value is None or value == '':
        return '0'
    
    try:
        # Convert to Decimal for precise handling
        dec_value = Decimal(str(value))
        
        # Check if value is zero
        if dec_value == 0:
            return '0'
        
        # Format with specified decimal places
        format_string = f"{{:.{decimals}f}}"
        formatted = format_string.format(float(dec_value))
        
        # Remove trailing zeros
        formatted = formatted.rstrip('0').rstrip('.')
        
        return formatted
            
    except (ValueError, TypeError):
        return str(value)


@register.filter
def percentage_display(value):
    """
    Display percentage values intelligently.
    Shows up to 4 decimal places for percentages.
    
    Usage: {{ value|percentage_display }}
    """
    if value is None or value == '':
        return '0'
    
    try:
        dec_value = Decimal(str(value))
        
        if dec_value == 0:
            return '0'
        
        # Convert to string
        str_value = str(dec_value)
        
        if '.' in str_value:
            integer_part, decimal_part = str_value.split('.')
            decimal_part = decimal_part.rstrip('0')[:4]  # Max 4 decimals for percentages
            
            if not decimal_part:
                return integer_part
            return f"{integer_part}.{decimal_part}"
        else:
            return str_value
            
    except (ValueError, TypeError):
        return str(value)
