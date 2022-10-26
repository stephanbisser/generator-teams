// Copyright (c) Wictor Wilén. All rights reserved.
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

import * as Generator from 'yeoman-generator';
import * as lodash from 'lodash';
import * as chalk from 'chalk';
import { GeneratorTeamsAppOptions } from './GeneratorTeamsAppOptions';
import { Yotilities } from './Yotilities';
import { ManifestGeneratorFactory } from './manifestGeneration/ManifestGeneratorFactory';
import * as inquirer from 'inquirer';
import { ManifestVersions } from './manifestGeneration/ManifestVersions';
import { v1 as uuid } from 'uuid';
import * as validate from 'uuid-validate';
import * as EmptyGuid from './EmptyGuid';
import * as crypto from 'crypto';
import { CoreFilesUpdaterFactory } from './coreFilesUpdater/CoreFilesUpdaterFactory';

const yosay = require('yosay');

// optimize App Insights performance
process.env.APPLICATION_INSIGHTS_NO_DIAGNOSTIC_CHANNEL = "none";
process.env.APPLICATION_INSIGHTS_NO_STATSBEAT = "true";
import * as AppInsights from 'applicationinsights';


let pkg: any = require('../../package.json');

/**
 * The main implementation for the `teams` generator
 */
export class GeneratorTeamsApp extends Generator {
    options: GeneratorTeamsAppOptions = new GeneratorTeamsAppOptions();

    public constructor(args: any, opts: any) {
        super(args, opts);
        opts.force = true;
        this.options.namespace = "teams";
        this.desc('Generate a Microsoft Teams application.');
        this.argument('solutionName', {
            description: 'Solution name, as well as folder name',
            required: false
        });
        this.argument('name', {
            description: 'Title of your Microsoft Teams App project',
            required: false
        });
        this.argument('developer', {
            description: 'Your (company) name',
            required: false
        });
        this.argument('manifestVersion', {
            description: 'The Teams manifest version you would like to use',
            required: false
        });
        this.argument('quickScaffolding', {
            description: 'Option to use quick scaffolding',
            required: false
        });
        this.argument('parts', {
            description: 'Determines which parts of the project to generate',
            required: false
        });
        this.argument('host', {
            description: 'The host URL of your solution',
            required: false
        });
        this.argument('showLoadingIndicator', {
            description: 'Option to show a loading indicator',
            required: false
        });
        this.argument('tabTitle', {
            description: 'The name of the tab to be set',
            required: false
        });
        this.argument('tabType', {
            description: 'The type of tab to generate',
            required: false
        });
        this.argument('tabScopes', {
            description: 'The scope of the tab to generate',
            required: false
        });
        this.argument('tabSSO', {
            description: 'Determines if you need SSO for your tab',
            required: false
        });
        this.option('telemetry', {
            type: Boolean,
            default: true,
            description: 'Pass usage telemetry, use --no-telemetry to not send telemetry. Note, no personal data is sent.'
        });
        // Set up telemetry
        if (this.options.telemetry &&
            !(process.env.YOTEAMS_TELEMETRY_OPTOUT === "1" ||
                process.env.YOTEAMS_TELEMETRY_OPTOUT === "true")) {


            // Set up the App Insights client
            const config = AppInsights.setup("6d773b93-ff70-45c5-907c-8edae9bf90eb");
            config.setInternalLogging(false, false);

            // Add a random session ID to the telemetry
            AppInsights.defaultClient.context.tags['ai.session.id'] = crypto.randomBytes(24).toString('base64');

            // Delete unnecessary telemetry data
            delete AppInsights.defaultClient.context.tags["ai.cloud.roleInstance"];
            delete AppInsights.defaultClient.context.tags["ai.cloud.role"];

            AppInsights.Configuration.setAutoCollectExceptions(true);
            AppInsights.Configuration.setAutoCollectPerformance(true);

            // Set common properties for all logging
            AppInsights.defaultClient.commonProperties = {
                version: pkg.version,
                node: process.version
            };

            AppInsights.defaultClient.trackEvent({ name: 'start-generator' });

        }

        this.options.existingManifest = this.fs.readJSON(`./src/manifest/manifest.json`);
    }

    public initializing() {
        this.log(yosay('Welcome to the ' + chalk.yellow(`Microsoft Teams App generator (${pkg.version})`)));
        this.composeWith('teams:tab', { 'options': this.options });
        this.composeWith('teams:bot', { 'options': this.options });
        this.composeWith('teams:custombot', { 'options': this.options });
        this.composeWith('teams:connector', { 'options': this.options });
        this.composeWith('teams:messageExtension', { 'options': this.options });
        this.composeWith('teams:localization', { 'options': this.options });

        // check schema version:
        const isSchemaVersionValid = ManifestGeneratorFactory.isSchemaVersionValid(this.options.existingManifest);
        if (!isSchemaVersionValid) {
            this.log(chalk.red('You are running the generator on an already existing project, but on a non supported-schema.'));
            if (this.options.telemetry) {
                AppInsights.defaultClient.trackEvent({ name: 'rerun-generator' });
                AppInsights.defaultClient.trackException({ exception: { name: 'Invalid schema', message: this.options.existingManifest["$schema"] } });
                AppInsights.defaultClient.flush();
            }
            process.exit(1);
        }
    }

    public prompting() {

        interface IAnswers {
            confirmedAdd: boolean;
            solutionName: string;
            whichFolder: string;
            name: string;
            developer: string;
            updateManifestVersion: boolean;
            manifestVersion: string;
            mpnId: string;
            parts: string[];
            host: string;
            unitTestsEnabled: boolean;
            lintingSupport: boolean;
            useAzureAppInsights: boolean;
            azureAppInsightsKey: string;
            updateBuildSystem: boolean;
            showLoadingIndicator: boolean;
            isFullScreen: boolean;
            quickScaffolding: boolean;
        };
        // find out what manifest versions we can use
        const manifestGeneratorFactory = new ManifestGeneratorFactory();
        const versions: inquirer.ChoiceOptions[] = ManifestGeneratorFactory.supportedManifestVersions.filter(version => {
            // filter out non supported upgrades
            if (this.options.existingManifest) {
                const manifestGenerator = manifestGeneratorFactory.createManifestGenerator(version.manifestVersion);
                return manifestGenerator.supportsUpdateManifest(this.options.existingManifest.manifestVersion);
            } else {
                return !version.hide; // always when not upgrading
            }
        }).map(version => {
            return {
                name: version.manifestVersion + (version.comment ? ` (${version.comment})` : ""),
                value: version.manifestVersion,
                extra: {
                    default: version.default
                }
            };
        })

        let generatorVersion = this.config.get("generator-version");
        if (!generatorVersion) {
            generatorVersion = "3.0.0";
        }

        const generatorPrefix = "[solution]";

        // return the question series
        return this.prompt<IAnswers>(
            [
                {
                    type: 'confirm',
                    name: 'confirmedAdd',
                    default: false,
                    message: `You are running the generator on an already existing project, "${this.options.existingManifest && this.options.existingManifest.name.short}", are you sure you want to continue?`,
                    prefix: generatorPrefix,
                    when: () => this.options.existingManifest,
                },
                {
                    type: 'confirm',
                    name: 'updateBuildSystem',
                    default: false,
                    message: 'Update yo teams core files? WARNING: Ensure your source code is under version control so you can merge any customizations of the core files!',
                    prefix: generatorPrefix,
                    when: (answers: IAnswers) => this.options.existingManifest && generatorVersion && generatorVersion != pkg.version && answers.confirmedAdd == true
                },
                {
                    type: 'input',
                    name: 'solutionName',
                    default: lodash.kebabCase(this.appname),
                    when: () => !(this.options.solutionName || this.options.existingManifest),
                    prefix: generatorPrefix,
                    message: 'What is your solution name?',
                },
                {
                    type: 'list',
                    name: 'whichFolder',
                    default: 'current',
                    when: () => !(this.options.solutionName || this.options.existingManifest),
                    message: 'Where do you want to place the files?',
                    prefix: generatorPrefix,
                    choices: [
                        {
                            name: 'Use the current folder',
                            value: 'current'
                        },
                        {
                            name: 'Create a sub folder with solution name',
                            value: 'subdir'
                        }
                    ]
                },
                {
                    type: 'input',
                    name: 'name',
                    message: 'Title of your Microsoft Teams App project?',
                    prefix: generatorPrefix,
                    when: () => !(this.options.name || this.options.existingManifest),
                    default: this.appname
                },
                {
                    type: 'input',
                    name: 'developer',
                    message: 'Your (company) name? (max 32 characters)',
                    prefix: generatorPrefix,
                    default: this.user.git.name,
                    validate: (input: string) => {
                        return input.length > 0 && input.length <= 32;
                    },
                    when: () => !(this.options.developer || this.options.existingManifest),
                    store: true
                },
                {
                    type: "confirm",
                    name: "updateManifestVersion",
                    message: `Do you want to change the current manifest version ${this.options.existingManifest && "(" + this.options.existingManifest.manifestVersion + ")"}?`,
                    prefix: generatorPrefix,
                    when: (answers: IAnswers) => this.options.existingManifest && versions.length > 0 && answers.confirmedAdd != false,
                    default: false
                },
                {
                    type: 'list',
                    name: 'manifestVersion',
                    message: 'Which manifest version would you like to use?',
                    prefix: generatorPrefix,
                    choices: versions,
                    default: versions.find((v: inquirer.ChoiceOptions) => v.extra.default) ?
                        versions.find((v: inquirer.ChoiceOptions) => v.extra.default)!.value :
                        (versions[0] ? versions[0].value : ""),
                    when: (answers: IAnswers) => (this.options.existingManifest && answers.updateManifestVersion && versions.length > 0) || !(this.options.manifestVersion || this.options.existingManifest)
                },
                {
                    type: "confirm",
                    name: "quickScaffolding",
                    message: `Quick scaffolding`,
                    prefix: generatorPrefix,
                    when: () => !(this.options.quickScaffolding || this.options.existingManifest),
                    default: true
                },
                {
                    type: 'input',
                    name: 'mpnId',
                    message: 'Enter your Microsoft Partner ID, if you have one? (Leave blank to skip)',
                    prefix: generatorPrefix,
                    default: undefined,
                    when: (answers: IAnswers) => !answers.quickScaffolding && !this.options.existingManifest,
                    validate: (input: string) => {
                        return input.length <= 10;
                    }
                },
                {
                    type: 'checkbox',
                    message: 'What features do you want to add to your project?',
                    prefix: generatorPrefix,
                    name: 'parts',
                    choices: [
                        {
                            name: 'A Tab',
                            value: 'tab',
                            checked: true,
                            disabled: () => {
                                if (this.options.existingManifest &&
                                    this.options.existingManifest.configurableTabs &&
                                    this.options.existingManifest.configurableTabs.length >= 1 &&
                                    this.options.existingManifest.staticTabs &&
                                    this.options.existingManifest.staticTabs.length >= 10) {
                                    // max 1 configurable tab and 10 static tabs allowed
                                    return true;
                                } else {
                                    return false;
                                }
                            },
                        },
                        {
                            name: 'A bot',
                            disabled: this.options.existingManifest,
                            value: 'bot'
                        },
                        {
                            name: 'An Outgoing Webhook',
                            disabled: this.options.existingManifest,
                            value: 'custombot'
                        },
                        {
                            name: 'A Connector (Not working with Teams JS SDK 2.0, please use version 3.5 of the generator)',
                            disabled: this.options.existingManifest,
                            value: 'connector'
                        },
                        {
                            name: 'A Message Extension Command',
                            disabled: () => {
                                if (this.options.existingManifest &&
                                    this.options.existingManifest.composeExtensions &&
                                    this.options.existingManifest.composeExtensions[0] &&
                                    this.options.existingManifest.composeExtensions[0].commands) {
                                    // max 10 commands are allowed
                                    return this.options.existingManifest.composeExtensions[0].commands.length >= 10;
                                } else {
                                    return false;
                                }
                            },
                            value: 'messageextension',
                        },
                        {
                            name: "Localization support",
                            value: "localization"
                        }
                    ],
                    when: (answers: IAnswers) => answers.confirmedAdd != false || !(this.options.tab)
                },
                {
                    type: 'input',
                    name: 'host',
                    message: 'The URL where you will host this solution?',
                    prefix: generatorPrefix,
                    default: (answers: IAnswers) => {
                        return `https://${lodash.camelCase(answers.solutionName).toLocaleLowerCase()}.azurewebsites.net`;
                    },
                    validate: Yotilities.validateUrl,
                    when: () => !this.options.existingManifest || !(this.options.host)
                },
                {
                    type: 'confirm',
                    name: 'showLoadingIndicator',
                    message: 'Would you like show a loading indicator when your app/tab loads?',
                    prefix: generatorPrefix,
                    default: false, // set to false until the 20 second timeout bug is fixed in Teams
                    when: () => !this.options.existingManifest || !(this.options.showLoadingIndicator)
                },
                {
                    type: 'confirm',
                    name: 'isFullScreen',
                    message: 'Would you like personal apps to be rendered without a tab header-bar?',
                    prefix: generatorPrefix,
                    default: false,
                    when: (answers: IAnswers) => !answers.quickScaffolding && !this.options.existingManifest,
                },
                {
                    type: 'confirm',
                    name: 'unitTestsEnabled',
                    message: 'Would you like to include Test framework and initial tests?',
                    prefix: generatorPrefix,
                    when: (answers: IAnswers) => !this.options.existingManifest && !answers.quickScaffolding,
                    store: true,
                    default: false
                },
                {
                    type: 'confirm',
                    name: 'lintingSupport',
                    message: 'Would you like to include ESLint support',
                    prefix: generatorPrefix,
                    when: (answers: IAnswers) => !this.options.existingManifest && !answers.quickScaffolding,
                    store: true,
                    default: true
                },
                {
                    type: 'confirm',
                    name: 'useAzureAppInsights',
                    message: 'Would you like to use Azure Applications Insights for telemetry?',
                    prefix: generatorPrefix,
                    when: (answers: IAnswers) => !this.options.existingManifest && !answers.quickScaffolding,
                    store: true,
                    default: false
                },
                {
                    type: 'input',
                    name: 'azureAppInsightsKey',
                    message: 'What is the Azure Application Insights Instrumentation Key?',
                    prefix: generatorPrefix,
                    default: (answers: IAnswers) => {
                        return EmptyGuid.empty;
                    },
                    validate: (input: string) => {
                        return validate(input) || input == EmptyGuid.empty;
                    },
                    when: (answers: IAnswers) => answers.useAzureAppInsights,
                },
            ]
        ).then((answers: IAnswers) => {
            if (answers.confirmedAdd == false) {
                process.exit(0);
            }
            if (!this.options.existingManifest) {
                // for new projects
                answers.host = answers.host.endsWith('/') ? answers.host.substr(0, answers.host.length - 1) : answers.host;
                this.options.title = answers.name;
                this.options.description = this.description;
                this.options.solutionName = this.options.solutionName || answers.solutionName;
                this.options.shouldUseSubDir = answers.whichFolder === 'subdir';
                this.options.libraryName = lodash.camelCase(this.options.solutionName);
                this.config.set("libraryName", this.options.libraryName);
                this.options.packageName = this.options.libraryName.toLocaleLowerCase();
                this.options.developer = answers.developer;
                this.options.host = answers.host;
                var tmp: string = this.options.host.substring(this.options.host.indexOf('://') + 3);
                this.options.hostname = this.options.host.substring(this.options.host.indexOf('://') + 3).toLocaleLowerCase();
                this.options.manifestVersion = answers.manifestVersion;
                this.options.mpnId = answers.mpnId;
                if (this.options.mpnId && this.options.mpnId.length == 0) {
                    this.options.mpnId = undefined;
                }
                var arr: string[] = tmp.split('.');
                this.options.namespace = lodash.reverse(arr).join('.').toLocaleLowerCase();
                this.options.id = uuid();
                if (this.options.host.indexOf('azurewebsites.net') >= 0) {
                    this.options.websitePrefix = this.options.host.substring(this.options.host.indexOf('://') + 3, this.options.host.indexOf('.'));
                } else {
                    this.options.websitePrefix = '[your Azure web app name]';
                }

                if (this.options.shouldUseSubDir) {
                    this.destinationRoot(this.destinationPath(this.options.solutionName));
                }
                this.options.showLoadingIndicator = answers.showLoadingIndicator;
                this.options.isFullScreen = answers.isFullScreen;
                this.options.unitTestsEnabled = answers.unitTestsEnabled;
                this.options.lintingSupport = answers.quickScaffolding || answers.lintingSupport;
                this.options.useAzureAppInsights = answers.useAzureAppInsights;
                this.options.azureAppInsightsKey = answers.azureAppInsightsKey;
            } else {
                // when updating projects
                this.options.developer = this.options.existingManifest.developer.name;
                this.options.title = this.options.existingManifest.name.short;
                this.options.hostname = "";
                this.options.useAzureAppInsights = this.config.get("useAzureAppInsights") || false;
                this.options.unitTestsEnabled = this.config.get("unitTestsEnabled") || false;
                let libraryName = Yotilities.getLibraryNameFromWebpackConfig(); // let's see if we can find the name in webpack.config.json (it might have been changed by the user)
                if (libraryName) {
                    this.options.libraryName = libraryName;
                } else {
                    // get the setting from Yo config
                    libraryName = this.config.get("libraryName");
                    if (libraryName) {
                        this.options.libraryName = libraryName!;
                    } else {
                        const pkg: any = this.fs.readJSON(`./package.json`);
                        this.log(chalk.yellow(`Unable to locate the library name in webpack.config.js, will use the package name instead (${pkg.name})`));
                        this.options.libraryName = pkg.name;
                    }
                }

                this.options.host = this.options.existingManifest.developer.websiteUrl;
                this.options.updateManifestVersion = answers.updateManifestVersion;
                this.options.manifestVersion = answers.manifestVersion ? answers.manifestVersion : ManifestGeneratorFactory.getManifestVersionFromValue(this.options.existingManifest.manifestVersion);
                this.options.updateBuildSystem = answers.updateBuildSystem;
            }

            if (answers.parts) {
                this.options.bot = (<string[]>answers.parts).indexOf('bot') != -1;
                this.options.tab = (<string[]>answers.parts).indexOf('tab') != -1;
                this.options.connector = (<string[]>answers.parts).indexOf('connector') != -1;
                this.options.customBot = (<string[]>answers.parts).indexOf('custombot') != -1;
                this.options.messageExtension = (<string[]>answers.parts).indexOf('messageextension') != -1;
                this.options.localization = (<string[]>answers.parts).indexOf('localization') != -1;
            }
            this.options.reactComponents = false; // set to false initially
        });
    }

    public configuring() {

    }

    public default() {

    }

    public writing() {
        this.sourceRoot();

        if (!this.options.existingManifest) {
            // This is a new project

            let staticFiles = [
                "_gitignore",
                "_vscode/launch.json",
                "src/server/tsconfig.json",
                "src/client/tsconfig.json",
                "src/manifest/icon-outline.png",
                "src/manifest/icon-color.png",
                "src/public/assets/icon.png",
                "src/public/styles/main.scss",
                "src/server/TeamsAppsComponents.ts",
                "Dockerfile"
            ];

            let templateFiles = [
                "README.md",
                "gulpfile.js",
                "package.json",
                ".env",
                'src/server/server.ts',
                "webpack.config.js",
                "src/client/client.ts",
                "src/public/index.html",
                "src/public/tou.html",
                "src/public/privacy.html",
            ];

            // Copy the manifest file with selected manifest version
            const manifestGeneratorFactory = new ManifestGeneratorFactory();
            const manifestGenerator = manifestGeneratorFactory.createManifestGenerator(this.options.manifestVersion);

            this.fs.writeJSON(
                Yotilities.fixFileNames("src/manifest/manifest.json", this.options),
                manifestGenerator.generateManifest(this.options)
            );

            templateFiles.forEach(t => {
                this.fs.copyTpl(
                    this.templatePath(t),
                    Yotilities.fixFileNames(t, this.options),
                    this.options);
            });

            // Add unit tests
            if (this.options.unitTestsEnabled) {
                staticFiles.push(
                    "src/test/test-setup.js",
                    "src/test/test-shim.js",
                    "src/client/jest.config.js",
                    "src/server/jest.config.js"
                );
                Yotilities.addAdditionalDevDeps([
                    ["enzyme", "^3.9.0"],
                    ["@types/enzyme", "^3.9.1"],
                    ["@types/jest", "^27.5.0"],
                    ["@types/enzyme-to-json", "^1.5.3"],
                    ["enzyme-adapter-react-16", "^1.11.2"],
                    ["enzyme-to-json", "^3.3.5"],
                    ["jest", "^28.1.0"],
                    ["ts-jest", "^28.0.2"],
                    ["jest-environment-jsdom", "^28.1.0"],
                    ["cheerio", "1.0.0-rc.10"]
                ], this.fs);

                Yotilities.addScript("test", "jest", this.fs);
                Yotilities.addScript("coverage", "jest --coverage", this.fs);

                Yotilities.addNode("jest", {
                    "projects": [
                        "src/client/jest.config.js",
                        "src/server/jest.config.js"
                    ]
                }, this.fs);
            }

            // add linting support
            if (this.options.lintingSupport) {
                staticFiles.push(
                    "_eslintignore",
                    "_eslintrc.json",
                    "src/server/.eslintrc.json",
                    "src/client/.eslintrc.json",
                );

                Yotilities.addAdditionalDevDeps([
                    ["@typescript-eslint/eslint-plugin", "^5.22.0"],
                    ["@typescript-eslint/parser", "^5.22.0"],
                    ["eslint", "^8.15.0"],
                    ["eslint-config-standard", "^17.0.0"],
                    ["eslint-plugin-import", "^2.22.1"],
                    ["eslint-plugin-node", "^11.1.0"],
                    ["eslint-plugin-promise", "^6.0.0"],
                    ["eslint-plugin-react", "^7.22.0"],
                    ["eslint-plugin-react-hooks", "^4.2.0"],
                    ["eslint-webpack-plugin", "^3.0.1"]
                ], this.fs);

                Yotilities.addScript("lint", "eslint ./src --ext .js,.jsx,.ts,.tsx", this.fs);
            }



            staticFiles.forEach(t => {
                this.fs.copy(
                    this.templatePath(t),
                    Yotilities.fixFileNames(t, this.options));
            });
        } else {
            if (this.options.updateBuildSystem) {
                let currentVersion = this.config.get("generator-version");
                if (!currentVersion) {
                    this.log(chalk.red("Nothing to update"));
                    if (this.options.telemetry) {
                        AppInsights.defaultClient.trackEvent({ name: 'update-core-files-empty', properties: { generatorVersion: pkg.version } });
                    }
                    process.exit(3);
                }

                const coreFilesUpdater = CoreFilesUpdaterFactory.createCoreFilesUpdater(currentVersion);
                if (coreFilesUpdater) {
                    const result = coreFilesUpdater.updateCoreFiles(this.options, this.fs, this.log);
                    if (this.options.telemetry) {
                        AppInsights.defaultClient.trackEvent({ name: 'update-core-files', properties: { result: result ? "true" : "false" } });
                    }
                    if (result === false) {
                        process.exit(4);
                    }
                } else {
                    this.log(chalk.red("WARNING: Unable to update build system automatically. See https://github.com/pnp/generator-teams/blob/master/docs/docs/upgrading-projects.md"));
                    if (this.options.telemetry) {
                        AppInsights.defaultClient.trackEvent({ name: 'update-core-files-failed', properties: { currentVersion, generatorVersion: pkg.version } });
                    }
                    process.exit(2);
                }
            }

            // running the generator on an already existing project
            if (this.options.updateManifestVersion) {
                const manifestGeneratorFactory = new ManifestGeneratorFactory();
                const manifestGenerator = manifestGeneratorFactory.createManifestGenerator(this.options.manifestVersion);
                this.fs.writeJSON(
                    Yotilities.fixFileNames("src/manifest/manifest.json", this.options),
                    manifestGenerator.updateManifest(this.options.existingManifest, this.log)
                );
            }
        }

        // if we have added any react based components
        if (this.options.reactComponents) {
            Yotilities.addAdditionalDeps([
                ["msteams-react-base-component", "^4.0.1"]
            ], this.fs);
        }

        if (this.options.useAzureAppInsights) {
            Yotilities.addAdditionalDeps([
                ["applicationinsights", "^1.3.1"]
            ], this.fs);
        }

        // Store the package version so that we can use it as a reference when upgrading
        this.config.set("generator-version", pkg.version);
    }

    public conflicts() {

    }

    public install() {
        // track usage
        if (this.options.telemetry) {
            if (this.options.existingManifest) {
                AppInsights.defaultClient.trackEvent({ name: 'rerun-generator' });
            }
            AppInsights.defaultClient.trackEvent({ name: 'end-generator' });
            if (this.options.bot) {
                AppInsights.defaultClient.trackEvent({
                    name: 'bot',
                    properties: {
                        type: this.options.botType,
                        files: this.options.botFilesEnabled ? "true" : "false",
                    }
                });
                if (this.options.botType == 'existing') {
                    AppInsights.defaultClient.trackEvent({ name: 'bot-existing' });
                } else {
                    AppInsights.defaultClient.trackEvent({ name: 'bot-new' });
                }
            }
            if (this.options.messageExtension) {
                AppInsights.defaultClient.trackEvent({
                    name: 'messageExtension',
                    properties: {
                        type: this.options.messagingExtensionType,
                        context: this.options.messagingExtensionActionContext ? this.options.messagingExtensionActionContext.join(";") : "",
                        input: this.options.messagingExtensionActionInputType || "",
                        response: this.options.messagingExtensionActionResponseType || "",
                        canUpdateConfiguration: this.options.messagingExtensionCanUpdateConfiguration ? "true" : "false",
                        inputConfig: this.options.messagingExtensionActionResponseTypeConfig ? "true" : "false"
                    }
                });
            }
            if (this.options.connector) {
                AppInsights.defaultClient.trackEvent({ name: 'connector' });
            }
            if (this.options.customBot) {
                AppInsights.defaultClient.trackEvent({ name: 'outgoingWebhook' });
            }
            if (this.options.staticTab) {
                AppInsights.defaultClient.trackEvent({ name: 'staticTab' });
            }
            if (this.options.tab) {
                AppInsights.defaultClient.trackEvent({ name: 'tab' });
            }
            if (this.options.unitTestsEnabled) {
                AppInsights.defaultClient.trackEvent({ name: 'unitTests' });
            }
            if (this.options.showLoadingIndicator) {
                AppInsights.defaultClient.trackEvent({ name: 'showLoadingIndicator' });
            }
            if (this.options.isFullScreen) {
                AppInsights.defaultClient.trackEvent({ name: 'isFullScreen' });
            }
            if (this.options.tabSSO) {
                AppInsights.defaultClient.trackEvent({ name: 'tabSSO' });
            }
            if (this.options.updateManifestVersion) {
                AppInsights.defaultClient.trackEvent({
                    name: 'updateManifest',
                    properties: {
                        from: this.options.existingManifest.manifestVersion,
                        to: this.options.manifestVersion
                    }
                });
            }
            if (this.options.localization) {
                AppInsights.defaultClient.trackEvent({
                    name: 'localization',
                    properties: {
                        defaultLanguage: this.options.defaultLanguage || "",
                        additionalLanguage: this.options.additionalLanguage || ""
                    }
                });
            }
            AppInsights.defaultClient.flush();
        }
    }

    public end() {
        this.log(chalk.yellow('Thanks for using the generator!'));
        this.log(chalk.yellow('Have fun and make great Microsoft Teams Apps...'));
    }
}
