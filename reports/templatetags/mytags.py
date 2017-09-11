from django import template
import datetime
import arrow
register = template.Library()

@register.filter
def change_date(value):
    b = arrow.get(value)
    a = b.strftime('%Y-%m-%d \b\b %H:%M:%S')
    return a