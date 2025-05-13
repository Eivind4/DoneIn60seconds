
// Loading the best practice rules developed by Michael Kovalsky to the location where TE is installed

// Download BPARules.json from GitHub to the correct folder based on TE version
var client = new System.Net.WebClient();

string localPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
string url = "https://raw.githubusercontent.com/microsoft/Analysis-Services/master/BestPracticeRules/BPARules.json";

// Default to TE2 path
string destination = System.IO.Path.Combine(localPath, @"TabularEditor\BPARules.json");

// Adjust for TE3 if needed
string version = typeof(ScriptHelper).Assembly.GetName().Version.Major.ToString();
if (version == "3")
{
    destination = System.IO.Path.Combine(localPath, @"TabularEditor3\BPARules.json");
}

// Download file
client.DownloadFile(url, destination);

// Confirm success
//Info("Best Practice Analyzer rules downloaded to:\n" + destination);
