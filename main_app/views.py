from django.http import HttpResponse

from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status

from .serializers import TrainingRecordSerializer
from .excel_utils import save_to_excel


class SubmitTrainingAPIView(APIView):
    """
    Resilient endpoint:
    - Validate payload first.
    - Try to save in DB.
    - Try to update/upload Excel to OneDrive only for a valid payload.
    - If DB is down, Excel can still be attempted using validated data.
    """

    def post(self, request, *args, **kwargs):
        training_id = None
        db_saved = False
        db_error = None

        # 1) Validate payload structure BEFORE any DB/Excel write.
        serializer = TrainingRecordSerializer(data=request.data)
        if not serializer.is_valid():
            return Response({
                "message": "Invalid payload",
                "db_saved": False,
                "training_id": None,
                "db_error": serializer.errors,
                "excel_saved": False,
                "excel_error": "Excel/OneDrive skipped because payload validation failed.",
            }, status=status.HTTP_400_BAD_REQUEST)

        # Make a plain validated payload for Excel/OneDrive.
        validated_payload = dict(serializer.validated_data)
        validated_payload["answers"] = [dict(a) for a in serializer.validated_data.get("answers", [])]

        # 2) Try saving to DB.
        try:
            training = serializer.save()
            training_id = training.id
            db_saved = True
        except Exception as e:
            db_saved = False
            db_error = str(e)

        # 3) Attempt Excel + OneDrive update only with validated data.
        excel_saved = False
        excel_error = None
        try:
            save_to_excel(validated_payload)
            excel_saved = True
        except Exception as e:
            excel_saved = False
            excel_error = str(e)

        # 4) Excel failure is reported explicitly; DB data remains available for recovery.
        if not excel_saved:
            return Response({
                "message": "Processed, but Excel/OneDrive failed",
                "db_saved": db_saved,
                "training_id": training_id,
                "db_error": db_error,
                "excel_saved": excel_saved,
                "excel_error": excel_error,
            }, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

        return Response({
            "message": "Processed successfully",
            "db_saved": db_saved,
            "training_id": training_id,
            "db_error": db_error,
            "excel_saved": excel_saved,
            "excel_error": excel_error,
        }, status=status.HTTP_201_CREATED)


def azure_callback(request):
    """
    Dummy endpoint for Azure App Registration.
    This endpoint is NOT used in authentication flow.
    """
    return HttpResponse("Azure callback endpoint is configured.")
