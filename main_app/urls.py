from django.urls import path
from .views import SubmitTrainingAPIView
from .views import test_large_upload_view


urlpatterns = [
    path('api/submit-training/', SubmitTrainingAPIView.as_view(), name='submit-training'),
    path("test-upload/", test_large_upload_view),

]
