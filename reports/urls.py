from django.conf.urls import url

from . import views

app_name = 'reports'
urlpatterns = [
    url(r'^$',views.index,name='index'),
    # url(r'^vote/$', views.vote, name='vote'),
    url(r'main/$',views.myview,name='main'),
    url(r'table/$',views.table,name='table'),
    url(r'^echo/$', views.echo_once),
    url(r'^test/$',views.test)
]
