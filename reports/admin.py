from django.contrib import admin
from .models import Config, Machine, Tag
# Register your models here.

admin.site.register(Config)
admin.site.register(Machine)
admin.site.register(Tag)
