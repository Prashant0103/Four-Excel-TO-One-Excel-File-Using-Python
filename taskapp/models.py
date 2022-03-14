from django.db import models
from django.db.models import fields
 
class data_sets(models.Model):
    excel_file = models.FileField()