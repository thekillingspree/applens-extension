const NOTE = "NOTE";

const DEFAULT_TEMPLATES = {
  note_fqr: {
    name: "FQR Note",
    id: "note_fqr",
    isDefault: true,
    template: `<p><strong><u>ISSUE</u></strong> </p><p><br></p><p><strong><u>App Details:</u></strong></p><ul><li><strong>App Name: </strong>{site}</li><li><strong>Subscription: </strong>{subscription}</li><li><strong>Resource Group: </strong>{resourceGroup}</li><li><strong>Issue Time (UTC): </strong>{date}</li><li><strong>Applens: </strong>{applens}</li><li><strong>Observer: </strong>{observer}</li><li><strong>ASC: </strong>{asc}</li></ul><p><br></p><p><strong><u>Troubleshooting:</u></strong> </p><p><br></p><p><br></p><p> </p>`,
  },
  note_follow: {
    name: "Follow-up Note",
    id: "note_follow",
    isDefault: true,
    template: `<p><strong><u>Status ({date})</u></strong> </p><p><br></p><p><br></p><p><strong><u>Next Action:</u></strong> </p><p><br></p><p><br></p><p><strong><u>Next Follow-up Date: ({next_followup_date})</u></strong> </p><p><br></p><p><br></p><p><strong>_____________________________________________</strong></p>`,
  },
  note_quick_ir: {
    name: "Quick IR Email",
    id: "note_quick_ir",
    isDefault: true,
    template: `<p>Hello **CUSTOMER NAME**, </p><p><br></p><p>Hope you are doing well! </p><p><br></p><p>Thank you for contacting Microsoft Support. My name is **YOUR NAME** and I'm from Azure App services team. I am the Support Professional who will be working with you on this Service Request #{caseNumber}. </p><p><br></p><p>I understand that ****cx issue****. Kindly share your availability for a call to discuss the issue further on a screen sharing session. Please let me know your time zone and the best time to reach you. </p><p><br></p><p>Looking forward to hearing back from you. </p><p><br></p><p>Thank you.</p>`,
  },
};

const TAB_COLORS = [
  "grey",
  "blue",
  "red",
  "yellow",
  "green",
  "pink",
  "purple",
  "cyan",
  "orange",
];

let groups = [
  {
    title: "Last 1 hour",
    id: "Last_1_hour",
    contexts: ["selection"],
  },
  {
    title: "Last 12 hours",
    id: "Last_12_hours",
    contexts: ["selection"],
  },
  {
    title: "Last 24 hours",
    id: "Last_24_hours",
    contexts: ["selection"],
  },
  {
    title: "Last 3 days",
    id: "Last_72_days",
    contexts: ["selection"],
  },
  {
    title: "Browse App",
    id: "Browse_App",
    contexts: ["selection"],
  },
  {
    title: "Open Observer",
    id: "observer",
    contexts: ["selection"],
  },
  {
    title: "Open ASC",
    id: "asc",
    contexts: ["all"],
    documentUrlPatterns: ["https://onesupport.crm.dynamics.com/*"],
  },
  {
    title: "Copy note/Email",
    id: "note",
    contexts: ["all"],
    documentUrlPatterns: ["https://onesupport.crm.dynamics.com/*"],
  },
  {
    title: "FQR Note",
    id: "note_fqr",
    parentId: "note",
    contexts: ["all"],
    documentUrlPatterns: ["https://onesupport.crm.dynamics.com/*"],
  },
  {
    title: "Follow-up Note",
    id: "note_follow",
    parentId: "note",
    contexts: ["all"],
    documentUrlPatterns: ["https://onesupport.crm.dynamics.com/*"],
  },
  {
    title: "Quick IR Email",
    id: "note_quick_ir",
    parentId: "note",
    contexts: ["all"],
    documentUrlPatterns: ["https://onesupport.crm.dynamics.com/*"],
  },
];

const addCustomTemplates = (templates) => {
  Object.values(templates).forEach((template) => {
    console.log(template);
    if (!template.isDefault) {
      groups.push({
        title: template.name,
        id: template.id,
        parentId: "note",
        contexts: ["all"],
        documentUrlPatterns: ["https://onesupport.crm.dynamics.com/*"],
      });
      createContextMenu();
    }
  });
};

chrome.runtime.onMessage.addListener(async (message) => {
  if (message.type === "sync") {
    const templates = await chrome.storage.sync.get("templates");
    if (Object.keys(templates).length > 0) {
      addCustomTemplates(templates);
    }
  }
});

chrome.runtime.onInstalled.addListener(async () => {
  const { templates } = await chrome.storage.sync.get(["templates"]);
  console.log(templates);
  addCustomTemplates(templates);
  if (!templates) {
    chrome.storage.sync.set({ templates: DEFAULT_TEMPLATES });
  }
});

const getDOMCaseNumber = (isNote = false) => {
  let caseNumber;
  try {
    caseNumber = document
      .querySelector('[id*="headerControlsList_"]')
      .innerText.split(" |")[0];

    if (!caseNumber) {
      throw new Error("Could not find case number. Opening prompt");
    }

    if (caseNumber.length > 18)
      caseNumber = caseNumber.slice(0, caseNumber.length - 3);
  } catch (error) {
    console.error(error);

    if (isNote) {
      alert("You are not on a case page. Please open a case and try again.");
      return;
    }

    caseNumber = window.prompt(
      "Looks like you are not in DFM page. Please Enter Case Number: "
    );
  }

  return caseNumber;
};

const getCaseNumber = (tab) => {
  return new Promise((resolve, reject) => {
    chrome.scripting.executeScript(
      {
        target: { tabId: tab.id },
        func: getDOMCaseNumber,
      },
      ([{ result: name }]) => (name ? resolve(name) : reject(name))
    );
  });
};

function getCurrentTime() {
  const isoStr = new Date().toISOString();
  return isoStr;
}

function getDate(startTime, difference, isMin = false) {
  let initDate = startTime;
  const dt = new Date(initDate);
  if (isMin) {
    dt.setMinutes(dt.getMinutes() - difference);
  } else {
    dt.setHours(dt.getHours() - difference);
  }

  return dt.toISOString();
}

const getApplensLinkForDate = async (appName, date, caseNumber) => {
  const endTime = getDate(date, 2);
  const startTime = getDate(date, 2);
  return `https://applens.trafficmanager.net/sites/${appName}?startTime=${startTime}&endTime=${endTime}&caseNumber=${caseNumber}`;
};

const getApplensLink = async (appName, time, caseNumber) => {
  const endTime = getDate(getCurrentTime(), 16, true);
  const startTime = getDate(endTime, time);
  return `https://applens.trafficmanager.net/sites/${appName}?startTime=${startTime}&endTime=${endTime}&caseNumber=${caseNumber}`;
};

const openApplens = async (appName, time, tab) => {
  const caseNumber = await getCaseNumber(tab);
  const applens = await getApplensLink(appName, time, caseNumber);
  createTab(applens, appName, caseNumber);
  //chrome.tabs.create({ url: appLens })
};

const openObserver = (appName) => {
  const observer = `https://wawsobserver.azurewebsites.windows.net/sites/${appName}`;
  //chrome.tabs.create({ url: observer })
  createTab(observer, appName);
};

const openASC = async (tab, selectedText) => {
  const caseNumber = await getCaseNumber(tab);

  const asc = `https://azuresupportcenter.msftcloudes.com/solutionexplorer?SourceId=OneSupport&srId=${caseNumber}`;
  //await chrome.tabs.create({ url: asc });
  createTab(asc, selectedText);
};

const browseApp = (name) => {
  //chrome.tabs.create({ url: `https://${name}.azurewebsites.net/` })
  createTab(`https://${name}.azurewebsites.net/`, name);
};

const createTab = async (url, appName, caseNumber = undefined) => {
  if (!appName && !caseNumber) {
    return chrome.tabs.create({ url });
  }

  const tab = await chrome.tabs.create({ url });
  const { grouping, groupingCaseNumber } = await chrome.storage.sync.get([
    "grouping",
    "groupingCaseNumber",
  ]);

  if (!grouping) {
    return;
  }

  let title = appName;

  if (groupingCaseNumber && caseNumber) title = caseNumber;

  const groups = await chrome.tabGroups.query({ title });

  if (groups.length > 0) {
    const groupId = groups[0].id;
    return await chrome.tabs.group({
      tabIds: tab.id,
      groupId,
    });
  }

  const groupId = await chrome.tabs.group({
    tabIds: tab.id,
  });
  const randomTabColor =
    TAB_COLORS[Math.floor(Math.random() * TAB_COLORS.length)];
  await chrome.tabGroups.update(groupId, {
    title,
    color: randomTabColor,
  });
};

const createContextMenu = () => {
  chrome.contextMenus.removeAll(function () {
    chrome.contextMenus.create({
      id: "app_lens",
      title: "Open in App Lens",
      contexts: ["all"],
    });

    groups.forEach((item) => {
      chrome.contextMenus.create({
        parentId: "app_lens",
        ...item,
      });
    });
  });
};

chrome.contextMenus.onClicked.addListener(async (data, tab) => {
  console.log(data);
  if (data.menuItemId === "Browse_App" && !!data.selectionText) {
    return browseApp(data.selectionText);
  }

  if (data.menuItemId === "observer" && !!data.selectionText) {
    return openObserver(data.selectionText);
  }

  if (data.menuItemId === "asc") {
    return openASC(tab, data.selectionText);
  }

  if (data.parentMenuItemId === "note") {
    return copyNote(tab, data.menuItemId);
  }

  openApplens(data.selectionText, parseInt(data.menuItemId.split("_")[1]), tab);
});

const copyNote = async (tab, note_id) => {
  chrome.tabs.sendMessage(tab.id, { type: NOTE, note_id });
};

createContextMenu();
setInterval(async () => {
  const templates = await chrome.storage.sync.get("templates");
  if (Object.keys(templates).length > 0) {
    addCustomTemplates(templates);
  }
}, 5000);
