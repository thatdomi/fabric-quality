
import json, requests
from azure.identity import DefaultAzureCredential, InteractiveBrowserCredential

class PowerBIRestHandler():
    def __init__(self):
        self.api = "https://analysis.windows.net/powerbi/api/.default"
        self.token = None
        self.powerBIBaseUrl = "https://api.powerbi.com/v1.0/myorg"
        

    def _authenticate(self):
        # Generates the access token via browser login
        auth = InteractiveBrowserCredential()
        # DefaultAzureCredential() could also be used
        self.token = auth.get_token(self.api).token
        
        if self.token:
            print("authenticated")

    def request_rest(self, url):
        # get token if none yet here
        if not self.token:
            self._authenticate()
        
        print(f"requesting: {url}")

        header = {'Authorization': f'Bearer {self.token}'}
        response = requests.get(url, headers=header)
        # Response code (200 = Success; 401 = Unauthorized; 404 = Bad Request)
        print(response)
        try:
            content = json.loads(response.content)
        except Exception as e:
            print(e)
            exit()
        return content

    def get_workspace_by_name(self, workspaceName) -> str:
        # returns the id of the provided worksapce
        workspaceName =workspaceName.replace(" ", "%20")
        url = f"{self.powerBIBaseUrl}/groups?$filter=contains(name,'{workspaceName}')&$top=1" #only get one entry back TODO: make more robust
        workspace = self.request_rest(url)
        workspaceId = workspace["value"][0]["id"]
        return workspaceId

    def get_reports_in_workspace(self, workspaceId) -> list:
        # returns the urls of all reports in workspace
        url = f"{self.powerBIBaseUrl}/groups/{workspaceId}/reports"
        reports = self.request_rest(url)
        report_urls = [ "https://app.powerbi.com/groups/" + workspaceId + "/reports/" + x["id"] for x in reports["value"]]
        return report_urls


if __name__ == "__main__":
    client = PowerBIRestHandler()
    workspaceId = client.get_workspace_by_name("Blog - Region")
    print(workspaceId)
    report_urls = client.get_reports_in_workspace(workspaceId=workspaceId)
    print(report_urls)