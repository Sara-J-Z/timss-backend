from django.core.management.base import BaseCommand
from main_app.services.graph_upload_session import GraphUploadSessionClient


class Command(BaseCommand):
    help = "Test Microsoft Graph large (chunked) upload session"

    def add_arguments(self, parser):
        parser.add_argument("--local", required=True, help="Local file path on Render")
        parser.add_argument("--folder", default="__TEST__", help="Remote folder under TIMSS")
        parser.add_argument("--name", default="test.xlsx", help="Remote filename")

    def handle(self, *args, **opts):
        client = GraphUploadSessionClient()
        result = client.upload_large_file(
            local_path=opts["local"],
            remote_folder=opts["folder"],
            remote_filename=opts["name"],
            chunk_size_mb=10
        )
        self.stdout.write(self.style.SUCCESS("✅ Upload completed"))
        self.stdout.write(str(result.get("webUrl", "")))
