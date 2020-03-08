<img src="https://ecprodpublic01.blob.core.windows.net/store/m_designs.jpg" title="M" alt="M" width="100" height="100">

# Microsoft Teams - Education
Microsoft Teams Automation

> Create Microsoft Teams with custom Channels, Tabs, Meetings etc using Microsoft Graph, PowerShell and Azure Automation Runbooks.

> Microsoft Teams, Azure Automation, PowerShell, Microsoft Graph API, Azure Labs, Education

## Microsoft References
- https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0 
- https://github.com/microsoftgraph/microsoft-graph-docs/tree/master/api-reference

## Design

### Source Data
Create a source JSON file to capture all the studentos, owners, units, classes and labs
Ask your friendly Student Information System Administrator to generate the JSON source file - or adapt to make calls directly to graph. 

```javascript
[
    {
        "Schedules": [],
        "Channels": [],
        "Owners": [
            ""
        ],
        "Members": [
            ""
        ],
        "Team": ""
    },
    {
        "Schedules": [
            {
                "Start_Time": "16:00",
                "Channel": "Prac-01-02",
                "Duration": "60",
                "Day": "Fri",
                "Location": "Social Sciences Lecture Theatre",
                "IsComputerLab": false,
                "Dates": "28/02/2020,06/03/2020,13/03/2020,20/03/2020,27/03/2020,03/04/2020,10/04/2020,24/04/2020,01/05/2020,08/05/2020,15/05/2020,22/05/2020"
            },
            {
                "Start_Time": "15:00",
                "Channel": "Prac-01-01",
                "Duration": "60",
                "Day": "Fri",
                "Location": "Social Sciences Lecture Theatre",
                "IsComputerLab": false,
                "Dates": "28/02/2020,06/03/2020,13/03/2020,20/03/2020,27/03/2020,03/04/2020,10/04/2020,24/04/2020,01/05/2020,08/05/2020,15/05/2020,22/05/2020"
            }],
        "Channels": [
            "Prac-01-02",
            "Prac-01-01",
        ],
        "Owners": [
            "Teacher1@school.edu.au",
            "Teacher2@school.edu.au"
        ],
        "Members": [
            "Student1@school.edu.au",
            "Student2@school.edu.au",
            "Student3@school.edu.au"
        ],
        "Team": "Maths 101 2020"
    }
]
```

### Processing 
Use Auzre Automation to schedule the creation and update of Microsoft Teams using PowerShell. 

You will need to create two App Registrations and a Teams Owner service account in Azure AD.
    App 1 - Will create the Office 365 Groups / Teams and manage the adding and removal of Teams, users and channels
    App 2 / Teams Owner - Will create the Teams Meetings for each channel

The Automation Powershell 
 ** Creates a new Team if it does not alredy exist. **
 ** Adds Channels **
 ** Adds Tabs to General Channel - Library or LMS etc **
 ** Adds users and owners to Group **



## Setup

