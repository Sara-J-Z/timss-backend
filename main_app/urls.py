from django.urls import path
from .views import SubmitTrainingAPIView
from .views import SubmitTrainingAPIView, azure_callback



urlpatterns = [
    path('api/submit-training/', SubmitTrainingAPIView.as_view(), name='submit-training'),
    path("auth/callback", azure_callback),  
]
