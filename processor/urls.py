from django.urls import path, re_path
from . import views

urlpatterns = [
    path('', views.home),
    path('upload/', views.upload_files),
    re_path(r'^download/(?P<download_id>.+)/$', views.download_file),
    path('download-all/', views.download_all),
    path('clear-all/', views.clear_all),
]