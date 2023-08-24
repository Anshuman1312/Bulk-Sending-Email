from django.urls import path
from .views import *
urlpatterns = [  
               
    path("Excel/",Excel.as_view()),
               
]