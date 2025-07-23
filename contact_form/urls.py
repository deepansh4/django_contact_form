from django.urls import path
from . import views

urlpatterns = [
    path('submit/', views.submit_contact_form, name='submit_contact_form'),
    path('submissions/', views.get_contact_submissions, name='get_contact_submissions'),

]
