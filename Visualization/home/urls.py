from django.contrib import admin
from django.urls import path
from home import views
from django.conf import settings
from django.conf.urls.static import static
# from .views import Index

urlpatterns = [

    path('', views.upload,name ="upload"),
    path('Visualization/',views.Visualization ,name = "Visualization"),

]
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL ,document_root = settings.MEDIA_ROOT)
