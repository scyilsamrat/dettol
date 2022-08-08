from django.contrib import admin

# Register your models here.
from django.contrib import admin
from .models import product,Invoice,Customer

admin.site.register(product)
admin.site.register(Invoice)
admin.site.register(Customer)