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

from django.http import JsonResponse
from main_app.services.graph_upload_session import GraphUploadSessionClient
import os


def test_large_upload_view(request):
    try:
        file_path = "/opt/render/project/src/excel_files/test.xlsx"

        if not os.path.exists(file_path):
            return JsonResponse({
                "error": "Test file not found",
                "path": file_path
            }, status=404)

        client = GraphUploadSessionClient()

        result = client.upload_large_file(
            local_path=file_path,
            remote_folder="__TEST__",
            remote_filename="test.xlsx",
        )

        return JsonResponse({
            "status": "success",
            "webUrl": result.get("webUrl")
        })

    except Exception as e:
        return JsonResponse({
            "status": "error",
            "message": str(e)
        }, status=500)
