from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from .serializers import TrainingRecordSerializer
from .excel_utils import save_to_excel


class SubmitTrainingAPIView(APIView):
    def post(self, request, *args, **kwargs):
        serializer = TrainingRecordSerializer(data=request.data)

        if serializer.is_valid():
            training = serializer.save()  # حفظ في PostgreSQL

            try:
                save_to_excel(serializer.data)
            except Exception as e:
                return Response({
                    "id": training.id,
                    "message": "Training record saved, but Excel upload failed",
                    "error": str(e)
                }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

            return Response({
                "id": training.id,
                "message": "Training record saved and uploaded successfully"
            }, status=status.HTTP_201_CREATED)

        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
