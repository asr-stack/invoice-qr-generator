from django.urls import path
from . import views

urlpatterns = [
    path('', views.home),
    path('upload/', views.upload_files),
    path('download/<str:filename>/', views.download_file),
]