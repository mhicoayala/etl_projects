"""
Google Sheets API interface.

This class serves as an interface to directly interact with Google sheets using their API.

Author: Mhico Ayala
Date: 2024-04-07
Changes: Class creation
"""

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from typing import Union

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
        self, spreadsheet_id: str, sheet_range: Union[str, dict]
    ) -> dict:
        """Get google sheet data.

        :param spreadsheet_id: Path or dict containing sheet credentials.
        :type spreadsheet_id: str, dict
        :param spreadsheet_id: Path or dict containing sheet credentials.
        :type spreadsheet_id: str, dict
        :return: Dictionary -- keys as column, values as rows
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
