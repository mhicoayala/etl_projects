"""
Google Sheets API interface.

This class serves as an interface to directly interact with Google sheets using their API.
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from typing import Union
import pandas as pd

scopes = ["https://www.googleapis.com/auth/spreadsheets"]


class GoogleSheetConnectorClass:
    """Work with Google Sheets using its API."""

    def __init__(self, credentials: Union[str, dict]) -> None:
        """Instantiate the GoogleSheetConnectorClass.

        :param credentials: Path or dict containing API credentials.
        :type credentials: str, dict
        :return: None
        :rtype: Nonetype
        """
        if isinstance(credentials, str):
            self.creds = Credentials.from_service_account_file(
                credentials, scopes=scopes
            )
        else:
            self.creds = Credentials.from_service_account_info(
                credentials, scopes=scopes
            )

    def build_connection(self) -> None:
        """Builds the connection to interact with API."""
        try:
            connection = build("sheets", "v4", credentials=self.creds)
            return connection

        except Exception as e:
            raise Exception(f"An error has occured: {e}")

    def get_sheet_data(
        self,
        spreadsheet_id: str,
        sheet_range: str,
    ) -> dict:
        """Get google sheet data.

        :param spreadsheet_id: Spreadsheet Id (found in the google sheet link).
        :type spreadsheet_id: str
        :param sheet_range: Sheet name + sheet range e.g.: (Sample Sheet!A:B).
        :type sheet_range: str
        :return: Dictionary -- keys as columns, values as rows
        :rtype: dict
        """

        # Establish connection with the API
        connection = self.build_connection()

        results = (
            connection.spreadsheets()  # type: ignore
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=sheet_range,
            )
        )

        try:
            results = results.execute()["values"]

            if not results:
                raise Exception("The spreadsheet/sheet range is empty.")

            else:
                column_names = results[0]
                results = {
                    column_names[i]: [row[i] for row in results[1:]]
                    for i in range(len(column_names))
                }
                return results

        except Exception as e:
            raise Exception(f"An error has occured: {e}")


def sheet_data_to_df(
    credentials: Union[str, dict], spreadsheet_id: str, sheet_range: str
) -> pd.DataFrame:
    """Transform sheet data to pandas dataframe.

    :param credentials: Path or dict containing API credentials.
    :type credentials: str, dict
    :param spreadsheet_id: Spreadsheet Id (found in the google sheet link).
    :type spreadsheet_id: str
    :param sheet_range: Sheet name + sheet range e.g.: (Sample Sheet!A:B).
    :type sheet_range: str
    :return: Pandas dataframe
    :rtype: pd.DataFrame
    """
    # Instatiate the class and build the connection
    gs = GoogleSheetConnectorClass(credentials=credentials)
    gs.build_connection()

    # Fetch the data from google sheet
    raw = gs.get_sheet_data(spreadsheet_id=spreadsheet_id, sheet_range=sheet_range)

    df = pd.DataFrame(data=raw)
    return df
