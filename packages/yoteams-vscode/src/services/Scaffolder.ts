import { parseWinPath } from './../utils/parseWinPath';
import { Executer } from './CommandExecuter';
import { Folders } from './Folders';
import { Notifications } from './Notifications';
import { Logger } from "./Logger";
import { commands, ProgressLocation, QuickPickItem, Uri, window } from 'vscode';
import { Commands, ComponentType, ComponentTypes, FrameworkTypes, ProjectFileContent } from '../constants';
import { Sample, Subscription } from '../models';
import { join } from 'path';
import { existsSync, mkdirSync, writeFileSync } from 'fs';
import * as glob from 'fast-glob';
import { ExtensionTypes } from '../constants/ExtensionTypes';
import { Extension } from './Extension';
import download from "github-directory-downloader/esm";
import { CliExecuter } from './CliCommandExecuter';
import { getPlatform } from '../utils';

export const PROJECT_FILE = 'project.pnp';

interface NameValue {
  name: string;
  value: string;
}

export class Scaffolder {

  public static registerCommands() {
    const subscriptions: Subscription[] = Extension.getInstance().subscriptions;
    
    subscriptions.push(
      commands.registerCommand(Commands.createProject, Scaffolder.createProject)
    ); 
  }
  
  /**
   * Create a new project
   * @returns 
   */
  public static async createProject() {
    Logger.info('Start creating a new project');

    const folderPath = await Scaffolder.getFolderPath();
    if (!folderPath) {
      Notifications.warning(`You must select the parent folder to create the project in`);
      return;
    }

    const solutionName = await Scaffolder.getSolutionName(folderPath);
    if (!solutionName) {
      Logger.warning(`Cancelled solution name input`);
      return;
    }

    Logger.info(`Creating a new project in ${folderPath}`);

    const yoCommand = `yo teams --skip-install`;
    Logger.info(`Command to execute: ${yoCommand}`);

    await window.withProgress({
      location: ProgressLocation.Notification,
      title: `Generating the new project... Check [output window](command:${Commands.showOutputChannel}) for more details`,
      cancellable: false
    }, async () => {
      try {
        const result = await Executer.executeCommand(folderPath, yoCommand);
        if (result !== 0) {
          Notifications.errorNoLog(`Failed to create the project. Check [output window](command:${Commands.showOutputChannel}) for more details.`);
          return;
        }

        Logger.info(`Command result: ${result}`);
        
        const newFolderPath = join(folderPath, solutionName);
        Scaffolder.createProjectFileAndOpen(newFolderPath, 'init');
      } catch (e) {
        Logger.error((e as Error).message);
        Notifications.errorNoLog(`Error creating the project. Check [output window](command:${Commands.showOutputChannel}) for more details.`);
      }
    });
  }


  /**
   * Create project file and open it in VS Code
   * @param folderPath 
   * @param content 
   */
  private static async createProjectFileAndOpen(folderPath: string, content: any) {
    writeFileSync(join(folderPath, PROJECT_FILE), content, { encoding: 'utf8' });

    if (getPlatform() === "windows") {
      await commands.executeCommand(`vscode.openFolder`, Uri.file(parseWinPath(folderPath)));
    } else {
      await commands.executeCommand(`vscode.openFolder`, Uri.parse(folderPath));
    }
  }

  /**
   * Get the name of the solution to create
   * @returns 
   */
  private static async getSolutionName(folderPath: string): Promise<string | undefined> {
    return await window.showInputBox({
      title: 'What is your solution name?',
      placeHolder: 'Enter your solution name',
      ignoreFocusOut: true,
      validateInput: (value) => {
        if (!value) {
          return 'Solution name is required';
        }

        const solutionPath = join(folderPath, value);
        if (existsSync(solutionPath)) {
          return `Folder with "${value}" already exists`;
        }

        return undefined;
      }
    });
  }

  /**
   * Select the path to create the project in
   * @returns 
   */
  private static async getFolderPath(): Promise<string | undefined> {
    const wsFolder = await Folders.getWorkspaceFolder();
    const folderOptions: QuickPickItem[] = [{
      label: '$(folder) Browse...',
      alwaysShow: true,
      description: 'Browse for the parent folder to create the project in'
    }];

    if (wsFolder) {
      folderOptions.push({
        label: `\$(folder-active) ${wsFolder.name}`,
        description: wsFolder.uri.fsPath
      });
    }

    const folderPath = await window.showQuickPick(folderOptions, {
      canPickMany: false,
      ignoreFocusOut: true,
      title: 'Select the parent folder to create the project in'
    }).then(async (selectedFolder) => {
      if (selectedFolder?.label === '$(folder) Browse...') {
        const folder = await window.showOpenDialog({
          canSelectFolders: true,
          canSelectFiles: false,
          canSelectMany: false,
          openLabel: 'Select',
          title: 'Select the parent folder where you want to create the project',
        });
        if (folder?.length) {
          return folder[0].fsPath;
        }
        return undefined;
      } else {
        return selectedFolder?.description;
      }
    });

    return folderPath;
  }
}