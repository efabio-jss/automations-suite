// This n8n workflow demonstrates how the AI Agent can function
// The Virtual Assistant queries predefined sources and available formats 
// It uses the OpenAI API to interpret inputs and provide responses

{
  "nodes": [
    {
      "parameters": {},
      "name": "Start",
      "type": "n8n-nodes-base.start",
      "typeVersion": 1,
      "position": [250, 300]
    },
    // Estrutura por País e Recursos
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "portugalData"
      },
      "name": "Process Portugal Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 200]
    },
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "spainData"
      },
      "name": "Process Spain Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 300]
 },
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "canadaData"
      },
      "name": "Process Canada Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 400]
 },
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "swedenData"
      },
      "name": "Process Sweden Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 500]
},
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "europeData"
      },
      "name": "Process Europe Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 600]
    },
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "worldData"
      },
      "name": "Process World Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 700]
},
    // Integrações Gerais Aplicáveis a Todos os Países e Dados
    {
      "parameters": {
        "operation": "listBoards"
      },
      "name": "Fetch Monday.com Data",
      "type": "n8n-nodes-base.mondayCom",
      "typeVersion": 1,
      "position": [1000, 200]
    },
    {
      "parameters": {
        "resource": "list",
        "operation": "getAll"
      },
      "name": "Fetch SharePoint Data",
      "type": "n8n-nodes-base.microsoftSharepoint",
      "typeVersion": 1,
      "position": [1000, 300]
 },
    {
      "parameters": {
        "teamId": "your-team-id",
        "channelId": "your-channel-id",
        "message": "AI Agent has processed new data."
      },
      "name": "Send Microsoft Teams Notification",
      "type": "n8n-nodes-base.microsoftTeams",
      "typeVersion": 1,
      "position": [1000, 400]
},
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "europeData"
      },
      "name": "Process Europe Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 600]
    },
    {
      "parameters": {
        "operation": "read",
        "dataPropertyName": "worldData"
      },
      "name": "Process World Data",
      "type": "n8n-nodes-base.function",
      "typeVersion": 1,
      "position": [750, 700]
},
    // Integrações Gerais Aplicáveis a Todos os Países e Dados
    {
      "parameters": {
        "operation": "listBoards"
      },
      "name": "Fetch Monday.com Data",
      "type": "n8n-nodes-base.mondayCom",
      "typeVersion": 1,
      "position": [1000, 200]
    },
    {
      "parameters": {
        "resource": "list",
        "operation": "getAll"
      },
      "name": "Fetch SharePoint Data",
      "type": "n8n-nodes-base.microsoftSharepoint",
      "typeVersion": 1,
      "position": [1000, 300]
 },
    {
      "parameters": {
        "teamId": "your-team-id",
        "channelId": "your-channel-id",
        "message": "AI Agent has processed new data."
      },
      "name": "Send Microsoft Teams Notification",
      "type": "n8n-nodes-base.microsoftTeams",
      "typeVersion": 1,
      "position": [1000, 400]

