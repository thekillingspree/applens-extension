class StoredTemplatesManager {
  constructor() {
    this.templates = null;
    this.loadTemplates();
  }

  loadTemplates() {
    setInterval(() => {
      chrome.storage.sync.get("templates", (data) => {
        this.templates = data.templates;
      });
    }, 5000);
  }

  getTemplate(templateId) {
    if (!this.templates) throw new Error("Templates not loaded");
    return this.templates[templateId];
  }

  async upsertTemplate(templateId, template) {
    this.templates[templateId] = template;
    await this.saveToStorage();
  }

  async saveToStorage() {
    chrome.storage.sync.set({ templates: this.templates });
  }

  async delete(templateId) {
    delete this.templates[templateId];
    await this.saveToStorage();
  }
}

class TemplateBuilder {
  constructor(template) {
    this.variables = new Map();
    this.template = template;
  }

  setTemplate(templateString) {
    this.template = templateString;
  }

  setVariable(variableName, variableValue, defaultValue = "N/A") {
    this.variables.set(variableName, variableValue || defaultValue);
  }

  setVariables(variables) {
    for (const [key, value] of Object.entries(variables)) {
      this.setVariable(key, value);
    }
  }

  compileTemplate() {
    const compiledTemplate = this.template.replace(
      /\{([^}]+)\}/g,
      (_, variableName) => {
        if (this.variables.has(variableName)) {
          return this.variables.get(variableName);
        } else {
          return `NOT FOUND`;
        }
      }
    );
    return compiledTemplate;
  }
}

class FQRTemplate extends TemplateBuilder {
  constructor() {
    const _template = templateManager.getTemplate("note_fqr")?.template;
    if (!_template) {
      throw new Error("Template not found");
    }
    super(_template);
  }

  getVerbatim() {
    const verbatim = getVerbatimDOM();
    const variables = getVariablesFromVerbatim(verbatim);

    this.setVariables(variables);
  }
}

class FollowUpTemplate extends TemplateBuilder {
  constructor() {
    const _template = templateManager.getTemplate("note_follow")?.template;
    if (!_template) {
      throw new Error("Template not found");
    }
    super(_template);

    this.variables.set("date", new Date().toLocaleDateString());
    const follow_up_date = this.getNextFollowUpDate();
    this.variables.set(
      "next_followup_date",
      follow_up_date.toLocaleDateString()
    );
  }

  getNextFollowUpDate() {
    //If weekend, next follow up is Monday, else the follow-up is 2 days from now
    const today = new Date();
    const day = today.getDay();
    const nextFollowUpDate = new Date();
    if (day > 0 && day < 5) {
      nextFollowUpDate.setDate(today.getDate() + 2);
    } else if (day === 5) {
      nextFollowUpDate.setDate(today.getDate() + 3);
    } else if (day === 6) {
      nextFollowUpDate.setDate(today.getDate() + 2);
    } else {
      nextFollowUpDate.setDate(today.getDate() + 1);
    }
    return nextFollowUpDate;
  }
}

class QuickIRTemplate extends TemplateBuilder {
  constructor() {
    const _template = templateManager.getTemplate("note_quick_ir")?.template;
    if (!_template) {
      throw new Error("Template not found");
    }
    super(_template);
    this.setVariable("caseNumber", getCaseNumber(), "**ENTER CASE NUMBER**");
  }
}

const NOTE = "NOTE";
const templateManager = new StoredTemplatesManager();

chrome.runtime.onMessage.addListener(({ type, ...others }) => {
  switch (type) {
    case NOTE:
      if (others.note_id) copyNote(others.note_id);
      break;
    default:
      console.error("Unknown message type", type);
      throw new Error("Unknown message type");
  }
});

function getDate(startTime, difference, isMin = false) {
  let initDate = startTime;
  const dt = new Date(`${initDate}.000Z`);
  if (isMin) {
    dt.setMinutes(dt.getMinutes() - difference);
  } else {
    dt.setHours(dt.getHours() - difference);
  }

  return dt.toISOString();
}

const getApplensLinkForDate = (appName, date, caseNumber) => {
  const endTime = getDate(date, -2);
  const startTime = getDate(date, 2);
  return `https://applens.trafficmanager.net/sites/${appName}?startTime=${startTime}&endTime=${endTime}&caseNumber=${caseNumber}`;
};

const getCaseNumber = () => {
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

    caseNumber = window.prompt(
      "Looks like you are not in DFM page. Please Enter Case Number: "
    );
  }

  console.log(caseNumber);
  return caseNumber;
};

function getAnchorTag(url, name) {
  return `<a href="${url}">${name || url}</a>`;
}

const getVariablesFromVerbatim = (verbatim) => {
  const observer = getAnchorTag(
    `https://wawsobserver.azurewebsites.windows.net/sites/${verbatim.site}`,
    "Observer"
  );

  const caseNumber = getCaseNumber();

  if (!caseNumber)
    return {
      ...verbatim,
      observer,
    };

  const applens = getAnchorTag(
    getApplensLinkForDate(verbatim.site, verbatim.date, caseNumber),
    "Applens"
  );

  const asc = getAnchorTag(
    `https://azuresupportcenter.msftcloudes.com/solutionexplorer?SourceId=OneSupport&srId=${caseNumber}`,
    "ASC"
  );

  return {
    ...verbatim,
    applens,
    observer,
    asc,
  };
};

const copyNote = (note_id) => {
  let template;
  switch (note_id) {
    case "note_fqr":
      template = new FQRTemplate();
      template.getVerbatim();
      break;
    case "note_follow":
      template = new FollowUpTemplate();
      break;
    case "note_quick_ir":
      template = new QuickIRTemplate();
      break;
    default:
      templateObject = templateManager.getTemplate(note_id);
      if (!templateObject) {
        console.error("Template not found");
        return;
      }
      console.log(template);
      template = new TemplateBuilder(templateObject.template);
      break;
  }

  const clipboardItem = new ClipboardItem({
    "text/html": new Blob([template.compileTemplate()], { type: "text/html" }),
  });

  navigator.clipboard.write([clipboardItem]);
};

const processVerbatim = (verbatim) => {
  let dateText = verbatim.match(/(ProblemStartTime):\s*(?<date>.+)/)?.groups
    ?.date;
  console.log(dateText);

  let site = verbatim.match(
    /(ResourceUri):\s*\/subscriptions\/(?<subscription>.+)\/resourceGroups\/(?<resourceGroup>.+)\/providers\/Microsoft.Web\/sites\/(?<site>.+)/
  )?.groups;

  if (!dateText) {
    console.warn("No Date found.");
    dateText = `${new Date()} (No Date found)`;
  }

  if (!site || !site.site) {
    console.warn("No Site found.");
    site = { site: "No Site found" };
  }

  return {
    date: dateText,
    ...site,
  };
};

const getVerbatimDOM = () => {
  const verbatimRAwText = document.querySelector(
    "[id*=customerstatement] textarea"
  ).value;

  const verbatim = processVerbatim(verbatimRAwText);

  if (!verbatim.site) {
    return {};
  }

  return verbatim;
};
