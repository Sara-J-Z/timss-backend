from django.urls import path
from .views import SubmitTrainingAPIView



urlpatterns = [
    path('api/submit-training/', SubmitTrainingAPIView.as_view(), name='submit-training'), 
]
