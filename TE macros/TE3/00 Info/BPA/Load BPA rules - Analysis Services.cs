
// Source for script: https://github.com/PowerBI-tips/TabularEditor-Scripts/blob/main/Basic/Import%20Best%20Practice%20Analyer%20Rules.csx


System.Net.WebClient w = new System.Net.WebClient(); 

string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
string url = "https://raw.githubusercontent.com/microsoft/Analysis-Services/master/BestPracticeRules/BPARules.json";
string version = System.Windows.Forms.Application.ProductVersion.Substring(0,1);
string downloadLoc = path+@"\TabularEditor\BPARules.json";

if (version == "3")
{
    downloadLoc = path+@"\TabularEditor3\BPARules.json";
}

w.DownloadFile(url, downloadLoc);