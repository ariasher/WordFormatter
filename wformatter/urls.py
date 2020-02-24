from django.urls import path
from wformatter import views


app_name = "wformatter"

urlpatterns = [
    path('/', views.index, name="index"),
    path("/<int:index>/" views.)
]
