"""dettol URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.2/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.urls import include,path
from django.contrib import admin
from main import views
from django.contrib import admin

admin.site.site_header = 'SCYIL PVT LTD'                    # default: "Django Administration"
admin.site.index_title = 'DETTOL SOFTWARE'                 # default: "Site administration"
admin.site.site_title = 'Scyil softwares'
urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.Home),
    path('addproduct/',views.Product),
    path('addcustomer/',views.customer),
    path('showcustomer/',views.showcustomer),
    path('print/',views.checker),
    path('showproduct/',views.showproduct),
    path('manageC/',views.manageC),
    path('manageP/',views.manageP),
    path('viewfile/',views.viewfile),

]
