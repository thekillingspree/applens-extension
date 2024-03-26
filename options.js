class StoredTemplatesManager {
  constructor() {
    this.templates = null;
    this.loadTemplates();
  }

  loadTemplates() {
    chrome.storage.sync.get("templates", (data) => {
      this.templates = data.templates;
    });
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

  size() {
    return Object.keys(this.templates).length;
  }
}

const templateManager = new StoredTemplatesManager();
let main;
let currentTemplate = null;
let deleteButton;

var quill = new Quill("#editor-container", {
  placeholder: "Your awesome template",
  theme: "snow",
});

// Save the template to storage when the save button is clicked
document.querySelector("#save-template").addEventListener("click", () => {
  if (!currentTemplate) return;

  currentTemplate.template = quill.root.innerHTML.replace("<p>&nbsp;</p>", "");
  templateManager.upsertTemplate(currentTemplate.id, currentTemplate);
  alert("Template saved");
  notifyBackground();
});

function setActiveTemplate(template) {
  currentTemplate = template;
  quill.setContents(null);
  document.querySelector(".template-name").innerText = template.name;
  const delta = quill.clipboard.convert(template.template);
  quill.setContents(delta);
  if (currentTemplate.isDefault) {
    deleteButton.classList.add("hide");
  } else {
    deleteButton.classList.remove("hide");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  // Load the templates from storage
  main = document.querySelector("main");
  loadTemplates();

  // Add a new template when the add template button is clicked
  const addButton = document.querySelector("#add-template");
  if (templateManager.size() === 8) {
    addButton.classList.add("hide");
  }
  addButton.addEventListener("click", async () => {
    const templateName = prompt("Enter a name for the new template:");
    if (templateName) {
      const template = {
        name: templateName,
        id: `${templateName}-${crypto.randomUUID()}`,
        isDefault: false,
        template: "",
      };

      await templateManager.upsertTemplate(template.id, template);
      notifyBackground();
      addTemplateToList(template);

      if (templateManager.size() === 8) {
        addButton.classList.add("hide");
      }
    }
  });

  deleteButton = document.querySelector("#delete-template");
  deleteButton.addEventListener("click", async () => {
    console.log("Clicked");
    if (!currentTemplate) return;

    const confirmDelete = confirm(
      `Are you sure you want to delete ${currentTemplate.name}?`
    );

    if (confirmDelete) {
      await templateManager.delete(currentTemplate.id);
      deleteTemplateFromList(currentTemplate.id);
      notifyBackground();
      currentTemplate = null;
      main.classList.add("hide");
      if (templateManager.size() < 8) {
        addButton.classList.remove("hide");
      }
    }
  });
});

const loadTemplates = () => {
  console.log("Loading templates", templateManager.templates);

  for (templateId in templateManager.templates) {
    const template = templateManager.getTemplate(templateId);
    console.log(template);
    addTemplateToList(template);
  }
};

const deleteTemplateFromList = (templateId) => {
  console.log("Deleteing template", templateId);
  document.querySelector(`#${templateId}`).remove();
};

const addTemplateToList = (template) => {
  const newTemplateItem = document.createElement("li");
  newTemplateItem.id = template.id;
  newTemplateItem.textContent = template.name;
  newTemplateItem.addEventListener("click", () => {
    main.classList.remove("hide");
    document
      .querySelector(".templates-list li.selected")
      ?.classList.remove("selected");
    newTemplateItem.classList.add("selected");
    setActiveTemplate(templateManager.getTemplate(template.id));
  });
  document
    .querySelector(".sidebar .templates-list")
    .appendChild(newTemplateItem);
};

const notifyBackground = () => {
  chrome.runtime.sendMessage({ type: "sync" });
};
