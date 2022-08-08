from django.db import models

# Create your models here.
from django.db import models
class product(models.Model):
    id = models.AutoField
    hsncode=models.CharField(max_length=50)
    pname = models.CharField(max_length=50)
    MRP =  models.FloatField(default=0)
    rate = models.FloatField(default=0)
    cgst = models.IntegerField(default=0)
    sgst = models.IntegerField(default=0)
    def __str__(self):
        return  self.pname


class Invoice(models.Model):
    id = models.AutoField
    partyname = models.CharField(max_length=100)
    def __str__(self):
        return  self.id

class Customer(models.Model):
    id = models.AutoField
    name = models.CharField(max_length=100)
    add = models.CharField(max_length=100)
    pno=models.CharField(max_length=11)
    gst=models.CharField(max_length=20)
    def __str__(self):
        return  self.name
