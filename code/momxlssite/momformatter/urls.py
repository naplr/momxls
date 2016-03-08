from django.conf.urls import patterns, url
import momformatter.views as views


urlpatterns = patterns('',
    #url(r'^test$', views.some_view, name='test'),
    url(r'^.*$', views.upload_file, name='index'),
)
