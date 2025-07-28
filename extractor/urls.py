# extractor/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path("", views.upload_pdf, name="upload_pdf"),
    path("api/extract/", views.extract_pdf_data, name="extract_pdf_api"),  
]