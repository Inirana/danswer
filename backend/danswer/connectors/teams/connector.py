from typing import Any

import msal
# ms graph and authenticator apis
from office365.graph_client import GraphClient  # type: ignore
from office365.onedrive.driveitems.driveItem import DriveItem  # type: ignore
from office365.onedrive.sites.site import Site


from danswer.connectors.interfaces import GenerateDocumentsOutput
from danswer.configs.app_configs import INDEX_BATCH_SIZE
from danswer.connectors.interfaces import LoadConnector
from danswer.connectors.interfaces import PollConnector
from danswer.connectors.interfaces import SecondsSinceUnixEpoch
from danswer.connectors.models import BasicExpertInfo
from danswer.connectors.models import ConnectorMissingCredentialError
from danswer.connectors.models import Document
from danswer.connectors.models import Section
from danswer.utils.logger import setup_logger


logger = setup_logger()


class TeamsConnector(PollConnector, LoadConnector):
    # need to have args that encompass
    # What docs the connector will process and where it will find those docs
    def __init__(
        self,
        batch_size: int = INDEX_BATCH_SIZE,
        messages: list[str] = [],
    ) -> None:
        self.batch_size = batch_size
        self.graph_client: GraphClient | None = None
        self.message_list: list[str] = messages

    # must maintain a dict of all required acces information
    def load_credentials(self, credentials: dict[str, Any]) -> dict[str, Any] | None:
        # aad == Azure Active Directory
        # maybe need a try/catch  here?
        aad_client_id = credentials["aad_client_id"]
        aad_client_secret = credentials["aad_client_secret"]
        aad_directory_id = credentials["aad_directory_id"]

        def _acquire_token_func() -> dict[str, Any]:
            """
            Acquire token via MSAL
            """
            authority_url = f"https://login.microsoftonline.com/{aad_directory_id}"
            app = msal.ConfidentialClientApplication(
                authority=authority_url,
                client_id=aad_client_id,
                client_credential=aad_client_secret,
            )
            token = app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )
            return token

        self.graph_client = GraphClient(_acquire_token_func)
        return None


        def get_all_messages(self):
            pass
        def _fetch_messages_from_graph(self):
            # not sure if I want to break this out or have it in get_all
            pass
        def extract_messages(self):
            pass
        def convert_message_to_document(self):
            pass
        def poll_source(self):
            pass


    if __name__ == "__main__":
        pass




