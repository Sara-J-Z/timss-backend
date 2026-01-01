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

            # حفظ نسخة في Excel
            try:
                excel_path = save_to_excel(serializer.data)
            except Exception as e:
                return Response({
                    "id": training.id,
                    "message": "Training record saved in DB, but failed to save Excel",
                    "error": str(e)
                }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

            return Response({
                "id": training.id,
                "message": "Training record saved successfully",
                "excel_path": excel_path
            }, status=status.HTTP_201_CREATED)
        else:
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

