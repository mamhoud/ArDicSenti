from django.conf .urls import url , include
from . import views


from myapp.views import download_file, addword

app_name = 'home'

urlpatterns = [

    url(r'^index/',views.index, name="index"),
    url(r'^download_file/', download_file , name="download"),
    url(r'^addword/', addword, name="addword"),
    url(r'^contribute/', views.contr_step_2, name="contribute"),
    url(r'^contrib_step1/', views.contr_step1, name="contrib_step1d"),
    url(r'^api-auth/', include('rest_framework.urls', namespace='rest_framework'))


]