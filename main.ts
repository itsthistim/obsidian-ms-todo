import {
	App,
	Plugin,
	PluginSettingTab,
	Setting,
	MarkdownPostProcessorContext,
	MarkdownRenderChild,
} from "obsidian";
// @ts-ignore
import axios from "axios";

interface MsTodoSettings {
	msAuthToken: string;
}

interface MsTodoTask {
	id: string;
	title: string;
	status: "notStarted" | "inProgress" | "completed";
	listId: string;
	listName: string;
	checklistItems?: MsTodoChecklistItem[];
}

interface MsTodoChecklistItem {
	id: string;
	displayName: string;
	isChecked: boolean;
}

const DEFAULT_SETTINGS: MsTodoSettings = {
	msAuthToken: "",
};

const GRAPH_API_BASE = "https://graph.microsoft.com/v1.0";

export default class MsTodoPlugin extends Plugin {
	settings: MsTodoSettings;

	async onload() {
		await this.loadSettings();
		this.addSettingTab(new MsTodoSettingTab(this.app, this));
		this.registerMarkdownCodeBlockProcessor("mstodo", this.processMsTodoBlock.bind(this));
	}

	onunload() {
		// Cleanup if needed
	}

	async loadSettings() {
		this.settings = Object.assign(
			{},
			DEFAULT_SETTINGS,
			await this.loadData()
		);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}

	private async processMsTodoBlock(
		source: string,
		el: HTMLElement,
		ctx: MarkdownPostProcessorContext
	) {
		// Create container for the tasks
		const container = el.createDiv({ cls: "mstodo-container" });

		// Show loading state
		container.setText("Loading MS Todo tasks...");

		try {
			// Fetch tasks from API
			const tasks = await this.fetchAllTasks();

			if (!tasks || tasks.length === 0) {
				container.setText(
					"No tasks found. Check your API token in settings."
				);
				return;
			}

			// Clear loading text
			container.empty();

			// Group tasks by list
			const tasksByList = this.groupTasksByList(tasks);

			// Render tasks
			for (const [listName, listTasks] of Object.entries(tasksByList)) {
				const listContainer = container.createDiv({
					cls: "mstodo-list",
				});
				listContainer.createEl("h4", {
					text: listName,
					cls: "mstodo-list-title",
				});

				const taskList = listContainer.createEl("ul", {
					cls: "mstodo-task-list",
				});

				for (const task of listTasks) {
					const taskItem = taskList.createEl("li", {
						cls: "mstodo-task-item",
					});

					// Create task content container
					const taskContent = taskItem.createDiv({
						cls: "mstodo-task-content",
					});

					// Create checkbox
					const checkbox = taskContent.createEl("input", {
						type: "checkbox",
						cls: "mstodo-checkbox",
					});
					checkbox.checked = task.status === "completed";
					checkbox.addEventListener("change", async () => {
						await this.updateTaskStatus(
							task.id,
							task.listId,
							checkbox.checked
						);
						// The checkbox state change will automatically update the UI
					});

					// Create task text
					const taskText = taskContent.createSpan({
						text: task.title,
						cls: `mstodo-task-text ${
							task.status === "completed" ? "completed" : ""
						}`,
					});

					// Render checklist items if they exist
					if (task.checklistItems && task.checklistItems.length > 0) {
						const checklistContainer = taskItem.createEl("ul", {
							cls: "mstodo-checklist",
						});

						for (const checklistItem of task.checklistItems) {
							const checklistItemEl = checklistContainer.createEl(
								"li",
								{
									cls: "mstodo-checklist-item",
								}
							);

							// Create checklist item content container
							const checklistContent = checklistItemEl.createDiv({
								cls: "mstodo-checklist-content",
							});

							// Create checklist item checkbox
							const checklistCheckbox = checklistContent.createEl(
								"input",
								{
									type: "checkbox",
									cls: "mstodo-checklist-checkbox",
								}
							);
							checklistCheckbox.checked = checklistItem.isChecked;
							checklistCheckbox.addEventListener(
								"change",
								async () => {
									await this.updateChecklistItemStatus(
										task.id,
										task.listId,
										checklistItem.id,
										checklistCheckbox.checked
									);
								}
							);

							// Create checklist item text
							const checklistText = checklistContent.createSpan({
								text: checklistItem.displayName,
								cls: `mstodo-checklist-text ${
									checklistItem.isChecked ? "completed" : ""
								}`,
							});
						}
					}
				}
			}
		} catch (error) {
			console.error("Error loading MS Todo tasks:", error);
			container.setText(`Error loading tasks: ${error.message}`);
		}
	}

	private async fetchAllTasks(): Promise<MsTodoTask[]> {
		if (!this.settings.msAuthToken) {
			throw new Error("MS Graph API token not configured");
		}

		try {
			// First, get all todo lists
			const listsResponse = await axios.get(
				`${GRAPH_API_BASE}/me/todo/lists`,
				{
					headers: {
						Authorization: `Bearer ${this.settings.msAuthToken}`,
						"Content-Type": "application/json",
					},
				}
			);

			const todoLists = listsResponse.data.value || [];
			const allTasks: MsTodoTask[] = [];

			// For each list, get its tasks
			for (const list of todoLists) {
				try {
					const tasksResponse = await axios.get(
						`${GRAPH_API_BASE}/me/todo/lists/${list.id}/tasks`,
						{
							headers: {
								Authorization: `Bearer ${this.settings.msAuthToken}`,
								"Content-Type": "application/json",
							},
						}
					);

					const tasks = tasksResponse.data.value || [];
					const tasksWithListInfo: MsTodoTask[] = [];

					// For each task, get its checklist items
					for (const task of tasks) {
						try {
							const checklistResponse = await axios.get(
								`${GRAPH_API_BASE}/me/todo/lists/${list.id}/tasks/${task.id}/checklistItems`,
								{
									headers: {
										Authorization: `Bearer ${this.settings.msAuthToken}`,
										"Content-Type": "application/json",
									},
								}
							);

							const checklistItems =
								checklistResponse.data.value || [];
							const formattedChecklistItems = checklistItems.map(
								(item: any) => ({
									id: item.id,
									displayName: item.displayName,
									isChecked: item.isChecked,
								})
							);

							tasksWithListInfo.push({
								id: task.id,
								title: task.title,
								status: task.status,
								listId: list.id,
								listName: list.displayName,
								checklistItems: formattedChecklistItems,
							});
						} catch (checklistError) {
							// If checklist fetch fails, still include the task without checklist
							console.warn(
								`Failed to fetch checklist for task ${task.title}:`,
								checklistError.message
							);
							tasksWithListInfo.push({
								id: task.id,
								title: task.title,
								status: task.status,
								listId: list.id,
								listName: list.displayName,
								checklistItems: [],
							});
						}
					}

					allTasks.push(...tasksWithListInfo);
				} catch (taskError) {
					console.warn(
						`Failed to fetch tasks for list ${list.displayName}:`,
						taskError.message
					);
				}
			}

			return allTasks;
		} catch (error: any) {
			console.error("Error fetching tasks:", error);

			if (error.response) {
				throw new Error(
					`Microsoft Graph API Error: ${
						error.response.data.error?.message ||
						"Unknown API error"
					}`
				);
			} else if (error.request) {
				throw new Error("Unable to connect to Microsoft Graph API");
			} else {
				throw new Error("An unexpected error occurred");
			}
		}
	}

	private async updateTaskStatus(
		taskId: string,
		listId: string,
		completed: boolean
	) {
		if (!this.settings.msAuthToken) {
			throw new Error("MS Graph API token not configured");
		}

		try {
			const newStatus = completed ? "completed" : "notStarted";

			await axios.patch(
				`${GRAPH_API_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
				{
					status: newStatus,
				},
				{
					headers: {
						Authorization: `Bearer ${this.settings.msAuthToken}`,
						"Content-Type": "application/json",
					},
				}
			);

			console.log(`Updated task ${taskId} to ${newStatus}`);
		} catch (error: any) {
			console.error("Error updating task status:", error);

			if (error.response) {
				throw new Error(
					`Microsoft Graph API Error: ${
						error.response.data.error?.message ||
						"Unknown API error"
					}`
				);
			} else if (error.request) {
				throw new Error("Unable to connect to Microsoft Graph API");
			} else {
				throw new Error("An unexpected error occurred");
			}
		}
	}

	private async updateChecklistItemStatus(
		taskId: string,
		listId: string,
		checklistItemId: string,
		isChecked: boolean
	) {
		if (!this.settings.msAuthToken) {
			throw new Error("MS Graph API token not configured");
		}

		try {
			await axios.patch(
				`${GRAPH_API_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
				{
					isChecked: isChecked,
				},
				{
					headers: {
						Authorization: `Bearer ${this.settings.msAuthToken}`,
						"Content-Type": "application/json",
					},
				}
			);

			console.log(
				`Updated checklist item ${checklistItemId} to ${
					isChecked ? "checked" : "unchecked"
				}`
			);
		} catch (error: any) {
			console.error("Error updating checklist item status:", error);

			if (error.response) {
				throw new Error(
					`Microsoft Graph API Error: ${
						error.response.data.error?.message ||
						"Unknown API error"
					}`
				);
			} else if (error.request) {
				throw new Error("Unable to connect to Microsoft Graph API");
			} else {
				throw new Error("An unexpected error occurred");
			}
		}
	}

	private groupTasksByList(
		tasks: MsTodoTask[]
	): Record<string, MsTodoTask[]> {
		const grouped: Record<string, MsTodoTask[]> = {};

		for (const task of tasks) {
			if (!grouped[task.listName]) {
				grouped[task.listName] = [];
			}
			grouped[task.listName].push(task);
		}

		return grouped;
	}
}

class MsTodoSettingTab extends PluginSettingTab {
	plugin: MsTodoPlugin;

	constructor(app: App, plugin: MsTodoPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		containerEl.createEl("h2", { text: "MS Todo Settings" });

		new Setting(containerEl)
			.setName("Microsoft Graph API Token")
			.setDesc(
				"Enter your Microsoft Graph API access token for To Do access"
			)
			.addText((text) =>
				text
					.setPlaceholder("Enter your MS Graph API token")
					.setValue(this.plugin.settings.msAuthToken)
					.onChange(async (value) => {
						this.plugin.settings.msAuthToken = value;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName("Get Access Token")
			.addButton((button) =>
				button
					.setButtonText("Open Guide")
					.setCta()
					.onClick(() => {
						// TODO: Open the guide file or show instructions
						console.log("Open token guide");
					})
			);
	}
}
