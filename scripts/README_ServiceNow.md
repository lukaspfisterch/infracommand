# ServiceNow Integration - Implementable Features

## 📋 ServiceNow-Search.ps1

### **✅ Already implemented:**
- ✅ ServiceNow-Search.ps1 script created
- ✅ Isolated browser sessions
- ✅ Variable resolution (${COMPUTERNAME}, ${USERNAME})
- ✅ Flexible search options (Global, Incident, Clipboard)

### **🚀 To be implemented:**

#### **1. InfraCommand Integration:**
```json
{
  "label": "ServiceNow Search",
  "type": "ps1",
  "path": "./scripts/ServiceNow-Search.ps1",
  "args": "-Instance 'https://your-instance.service-now.com' -Query 'host=${COMPUTERNAME} OR user_name=${USERNAME}'",
  "elevate": false,
  "risk": "read",
  "weight": 30
}
```

#### **2. Various ServiceNow Buttons:**
- **ServiceNow Global Search** - General search
- **ServiceNow Incidents** - Incident-specific search
- **ServiceNow My Tickets** - Personal tickets
- **ServiceNow Last 7 Days** - Last 7 days

#### **3. Advanced Search Parameters:**
- **Host-specific search:** `host=${COMPUTERNAME}`
- **User-specific search:** `user_name=${USERNAME}`
- **Time range filter:** `sys_created_on>javascript:gs.daysAgo(7)`
- **Status filter:** `state!=closed`

#### **4. Browser Options:**
- **Edge** (Default) - `msedge.exe`
- **Chrome** - `chrome.exe`
- **Firefox** - `firefox.exe`

#### **5. Clipboard Integration:**
- **FromClipboard Parameter** - Automatically from clipboard
- **Interactive Mode** - Prompt when query is missing

### **💡 Usage:**

#### **Global Search:**
```powershell
.\ServiceNow-Search.ps1 -Query "host=COMPUTER01 OR user_name=admin"
```

#### **Incident Search:**
```powershell
.\ServiceNow-Search.ps1 -IncidentNumber "INC123456"
```

#### **From Clipboard:**
```powershell
.\ServiceNow-Search.ps1 -FromClipboard
```

#### **Interactive:**
```powershell
.\ServiceNow-Search.ps1
```

### **🔧 Technical Details:**
- **Isolated Sessions:** Each call creates temporary browser profile
- **Window Placement:** Helps InfraCommand with window detection
- **SSO-Avoidance:** Prevents tab reuse and SSO issues
- **Variable-Substitution:** PowerShell resolves ${COMPUTERNAME} and ${USERNAME}

### **📊 Status:**
- **Script:** ✅ Complete
- **Integration:** ⏳ Pending
- **Testing:** ⏳ Pending
- **Deployment:** ⏳ Pending
