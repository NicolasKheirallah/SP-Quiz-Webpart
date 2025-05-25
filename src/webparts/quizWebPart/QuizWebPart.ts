import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'QuizWebPartStrings';
import Quiz from './components/Quiz';
import { IQuizProps } from './components/IQuizProps';
import { IQuizPropertyPaneProps, IQuizQuestion, QuestionType } from './components/interfaces';
import QuizPropertyPane from './components/QuizPropertyPane';

// Replace the IQuizWebPartProps interface in QuizWebPart.ts with this updated version
export interface IQuizWebPartProps {
  title: string;
  questionsPerPage: number;
  successMessage: string;
  excellentScoreMessage: string;
  goodScoreMessage: string;
  averageScoreMessage: string;
  poorScoreMessage: string;
  errorMessage: string;
  resultsSavedMessage: string;
  showProgressIndicator: boolean;
  randomizeQuestions: boolean;
  randomizeAnswers: boolean;
  questions: IQuizQuestion[];
  passingScore: number;
  timeLimit: number;
  enableQuestionTimeLimit: boolean;
  defaultQuestionTimeLimit: number;
  resultsListName: string;
  
  // HTTP Trigger properties
  enableHttpTrigger: boolean;
  httpTriggerUrl: string;
  httpTriggerScoreThreshold: number;
  httpTriggerMethod: string;
  httpTriggerIncludeUserData: boolean;
  httpTriggerIncludeQuizData: boolean;
  httpTriggerCustomHeaders: string;
  httpTriggerTimeout: number;
}

export default class QuizWebPart extends BaseClientSideWebPart<IQuizWebPartProps> {
  private _reactElement: HTMLElement | null = null;
  private _propertyPaneContainer: HTMLElement | null = null;
  private _listOptions: { key: string; text: string }[] = [];
  private _listsLoaded: boolean = false;

// Update the render method in QuizWebPart.ts to pass HTTP trigger props
public render(): void {
  // Unmount any existing React components
  this.disposeReactComponents();

  // Create a new container for the main React component
  this._reactElement = document.createElement('div');
  this.domElement.appendChild(this._reactElement);

  const element: React.ReactElement<IQuizProps> = React.createElement(
    Quiz,
    {
      title: this.properties.title || 'SharePoint Quiz',
      questionsPerPage: this.properties.questionsPerPage || 5,
      context: this.context,
      displayMode: this.displayMode,
      updateProperty: (value: string) => {
        this.properties.title = value;
      },
      successMessage: this.properties.successMessage || 'Your score has been successfully recorded!',
      excellentScoreMessage: this.properties.excellentScoreMessage || 'Excellent! You have mastered this topic!',
      goodScoreMessage: this.properties.goodScoreMessage || 'Good job! You have a solid understanding.',
      averageScoreMessage: this.properties.averageScoreMessage || 'Not bad. There\'s room for improvement.',
      poorScoreMessage: this.properties.poorScoreMessage || 'Keep studying. You\'ll get better with practice.',
      errorMessage: this.properties.errorMessage || 'An error occurred. Please try again later.',
      resultsSavedMessage: this.properties.resultsSavedMessage || 'Your score has been successfully saved!',
      showProgressIndicator: this.properties.showProgressIndicator !== undefined ? this.properties.showProgressIndicator : true,
      randomizeQuestions: this.properties.randomizeQuestions !== undefined ? this.properties.randomizeQuestions : false,
      randomizeAnswers: this.properties.randomizeAnswers !== undefined ? this.properties.randomizeAnswers : false,
      passingScore: this.properties.passingScore || 70,
      timeLimit: this.properties.timeLimit ? this.properties.timeLimit * 60 : undefined,
      enableQuestionTimeLimit: this.properties.enableQuestionTimeLimit || false,
      defaultQuestionTimeLimit: this.properties.defaultQuestionTimeLimit || 60,
      questions: this.properties.questions || [],
      resultsListName: this.properties.resultsListName || 'QuizResults',
      
      // NEW: HTTP Trigger properties
      enableHttpTrigger: this.properties.enableHttpTrigger || false,
      httpTriggerUrl: this.properties.httpTriggerUrl || '',
      httpTriggerScoreThreshold: this.properties.httpTriggerScoreThreshold || 80,
      httpTriggerMethod: this.properties.httpTriggerMethod || 'POST',
      httpTriggerIncludeUserData: this.properties.httpTriggerIncludeUserData !== false,
      httpTriggerIncludeQuizData: this.properties.httpTriggerIncludeQuizData !== false,
      httpTriggerCustomHeaders: this.properties.httpTriggerCustomHeaders || '',
      httpTriggerTimeout: this.properties.httpTriggerTimeout || 30,
      
      updateQuestions: (questions: IQuizQuestion[]) => {
        this.properties.questions = questions;
        this.render();
      }
    }
  );

  // Render the new component
  ReactDom.render(element, this._reactElement);

  // Render property pane content if it's open
  this.renderPropertyPaneContent();
}

  // Helper function to clean up React components
  private disposeReactComponents(): void {
    // Unmount main React component
    if (this._reactElement) {
      ReactDom.unmountComponentAtNode(this._reactElement);
      if (this._reactElement.parentNode) {
        this._reactElement.parentNode.removeChild(this._reactElement);
      }
      this._reactElement = null;
    }

    // Unmount property pane React component
    if (this._propertyPaneContainer) {
      ReactDom.unmountComponentAtNode(this._propertyPaneContainer);
      if (this._propertyPaneContainer.parentNode) {
        this._propertyPaneContainer.parentNode.removeChild(this._propertyPaneContainer);
      }
      this._propertyPaneContainer = null;
    }
  }

  // Render the property pane content
  private renderPropertyPaneContent(): void {
    if (this.context.propertyPane.isPropertyPaneOpen()) {
      try {
        // Ensure proper container exists
        if (!this._propertyPaneContainer) {
          this._propertyPaneContainer = document.createElement('div');
          this._propertyPaneContainer.className = 'quiz-property-pane-container';
        }

        // Create the property pane component with error handling
        const propertyPaneElement = React.createElement<IQuizPropertyPaneProps>(
          QuizPropertyPane,
          {
            questions: this.properties.questions || [],
            onUpdateQuestions: (questions: IQuizQuestion[]) => {
              this.properties.questions = questions;
              this.render();
            },
            context: this.context  // Pass context to property pane
          }
        );

        // Render the new property pane component
        ReactDom.render(propertyPaneElement, this._propertyPaneContainer);
      } catch (error) {
        console.error('Error rendering property pane content:', error);
        // Create a simple error message if rendering fails
        if (this._propertyPaneContainer) {
          this._propertyPaneContainer.innerHTML = '<div style="color: red; padding: 10px;">Error loading question manager. Please refresh the page and try again.</div>';
        }
      }
    }
  }

  protected onDispose(): void {
    // Ensure we unmount all React components to prevent memory leaks
    this.disposeReactComponents();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  // Method to fetch available lists
  private async _getLists(): Promise<void> {
    if (!this._listsLoaded) {
      try {
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Title`,
          SPHttpClient.configurations.v1
        );

        if (response.ok) {
          const data = await response.json();
          this._listOptions = data.value.map((list: { Title: string }) => ({
            key: list.Title,
            text: list.Title
          }));
          
          // Always add QuizResults as a default option if it doesn't exist
          if (!this._listOptions.some(option => option.key === 'QuizResults')) {
            this._listOptions.unshift({
              key: 'QuizResults',
              text: 'QuizResults (Default)'
            });
          }
          
          this._listsLoaded = true;
          this.context.propertyPane.refresh();
        } else {
          console.error('Error fetching lists:', response.statusText);
          this._listOptions = [
            { key: 'QuizResults', text: 'QuizResults (Default)' }
          ];
        }
      } catch (error) {
        console.error('Error in _getLists method:', error);
        this._listOptions = [
          { key: 'QuizResults', text: 'QuizResults (Default)' }
        ];
      }
    }
  }

  // Initialize default questions if none exist
  protected onPropertyPaneConfigurationStart(): void {
    // Initialize with sample questions if none exist
    if (!this.properties.questions || this.properties.questions.length === 0) {
      this.properties.questions = [
        {
          id: 1,
          title: 'What is SharePoint?',
          category: 'SharePoint Basics',
          type: QuestionType.MultipleChoice,
          choices: [
            { id: '1', text: 'A content management system', isCorrect: true },
            { id: '2', text: 'A programming language', isCorrect: false },
            { id: '3', text: 'A hardware device', isCorrect: false },
            { id: '4', text: 'A database system', isCorrect: false }
          ]
        },
        {
          id: 2,
          title: 'SharePoint is developed by which company?',
          category: 'SharePoint Basics',
          type: QuestionType.MultipleChoice,
          choices: [
            { id: '1', text: 'Google', isCorrect: false },
            { id: '2', text: 'Microsoft', isCorrect: true },
            { id: '3', text: 'Apple', isCorrect: false },
            { id: '4', text: 'Amazon', isCorrect: false }
          ]
        }
      ];
    }

    // Set default results list name if not set
    if (!this.properties.resultsListName) {
      this.properties.resultsListName = 'QuizResults';
    }

    // Clean up any existing components
    this.disposeReactComponents();

    // Create the container for the property pane component
    this._propertyPaneContainer = document.createElement('div');
    this._propertyPaneContainer.className = 'quiz-property-pane-container';

    // Load lists for the dropdown
    this._getLists().catch(error => {
      console.error('Error loading lists:', error);
    });
  }

  protected onPropertyPaneConfigurationEnd(): void {
    // Clean up any property pane components
    this.disposeReactComponents();
  }

  // Handle property pane changes
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    // If questions were changed, trigger re-render
    if (propertyPath === 'questions' || propertyPath === 'resultsListName') {
      this.render();
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

// Update the getPropertyPaneConfiguration method in QuizWebPart.ts
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField('title', {
                label: strings.TitleFieldLabel
              }),
              PropertyPaneDropdown('questionsPerPage', {
                label: strings.QuestionsPerPageFieldLabel,
                options: [
                  { key: 1, text: '1' },
                  { key: 3, text: '3' },
                  { key: 5, text: '5' },
                  { key: 10, text: '10' }
                ]
              }),
              PropertyPaneToggle('showProgressIndicator', {
                label: 'Show Progress Indicator',
                onText: 'On',
                offText: 'Off'
              }),
              PropertyPaneToggle('randomizeQuestions', {
                label: 'Randomize Questions',
                onText: 'On',
                offText: 'Off'
              }),
              PropertyPaneToggle('randomizeAnswers', {
                label: 'Randomize Answer Choices',
                onText: 'On',
                offText: 'Off'
              })
            ]
          },
          {
            groupName: 'Quiz Settings',
            groupFields: [
              PropertyPaneTextField('passingScore', {
                label: 'Passing Score (%)',
                description: 'Minimum percentage to pass the quiz'
              }),
              PropertyPaneTextField('timeLimit', {
                label: 'Overall Quiz Time Limit (minutes)',
                description: 'Maximum time allowed for the entire quiz (0 for no time limit)'
              }),
              PropertyPaneToggle('enableQuestionTimeLimit', {
                label: 'Enable Question Time Limits',
                onText: 'On',
                offText: 'Off',
                checked: this.properties.enableQuestionTimeLimit
              }),
              this.properties.enableQuestionTimeLimit ? 
              PropertyPaneSlider('defaultQuestionTimeLimit', {
                label: 'Default Question Time Limit (seconds)',
                min: 10,
                max: 300,
                step: 5,
                showValue: true,
                value: this.properties.defaultQuestionTimeLimit || 60
              }) : null,
              PropertyPaneDropdown('resultsListName', {
                label: 'Quiz Results List',
                options: this._listOptions,
                selectedKey: this.properties.resultsListName || 'QuizResults'
              })
            ].filter(Boolean) as IPropertyPaneField<IPropertyPaneCustomFieldProps>[]
          },
          {
            groupName: 'HTTP Trigger Settings',
            groupFields: [
              PropertyPaneToggle('enableHttpTrigger', {
                label: 'Enable HTTP Trigger',
                onText: 'Enabled',
                offText: 'Disabled',
                checked: this.properties.enableHttpTrigger || false
              }),
              this.properties.enableHttpTrigger ? PropertyPaneTextField('httpTriggerUrl', {
                label: 'HTTP Trigger URL',
                description: 'Flow or Logic App HTTP trigger URL',
                placeholder: 'https://prod-xx.westus.logic.azure.com:443/workflows/...',
                multiline: false
              }) : null,
              this.properties.enableHttpTrigger ? PropertyPaneSlider('httpTriggerScoreThreshold', {
                label: 'Score Threshold (%)',
                min: 0,
                max: 100,
                step: 5,
                showValue: true,
                value: this.properties.httpTriggerScoreThreshold || 80
              }) : null,
              this.properties.enableHttpTrigger ? PropertyPaneDropdown('httpTriggerMethod', {
                label: 'HTTP Method',
                options: [
                  { key: 'POST', text: 'POST' },
                  { key: 'PUT', text: 'PUT' },
                  { key: 'PATCH', text: 'PATCH' },
                  { key: 'GET', text: 'GET' }
                ],
                selectedKey: this.properties.httpTriggerMethod || 'POST'
              }) : null,
              this.properties.enableHttpTrigger ? PropertyPaneToggle('httpTriggerIncludeUserData', {
                label: 'Include User Data',
                onText: 'Yes',
                offText: 'No',
                checked: this.properties.httpTriggerIncludeUserData !== false
              }) : null,
              this.properties.enableHttpTrigger ? PropertyPaneToggle('httpTriggerIncludeQuizData', {
                label: 'Include Quiz Results Data',
                onText: 'Yes',
                offText: 'No',
                checked: this.properties.httpTriggerIncludeQuizData !== false
              }) : null,
              this.properties.enableHttpTrigger ? PropertyPaneTextField('httpTriggerCustomHeaders', {
                label: 'Custom Headers (JSON)',
                description: 'Optional custom headers as JSON object: {"Authorization": "Bearer token"}',
                multiline: true,
                placeholder: '{"Content-Type": "application/json"}'
              }) : null,
              this.properties.enableHttpTrigger ? PropertyPaneSlider('httpTriggerTimeout', {
                label: 'Request Timeout (seconds)',
                min: 5,
                max: 60,
                step: 5,
                showValue: true,
                value: this.properties.httpTriggerTimeout || 30
              }) : null
            ].filter(Boolean) as IPropertyPaneField<IPropertyPaneCustomFieldProps>[]
          },
          {
            groupName: 'Messages',
            groupFields: [
              PropertyPaneTextField('successMessage', {
                label: 'Success Message'
              }),
              PropertyPaneTextField('excellentScoreMessage', {
                label: 'Excellent Score Message (90-100%)'
              }),
              PropertyPaneTextField('goodScoreMessage', {
                label: 'Good Score Message (70-89%)'
              }),
              PropertyPaneTextField('averageScoreMessage', {
                label: 'Average Score Message (50-69%)'
              }),
              PropertyPaneTextField('poorScoreMessage', {
                label: 'Poor Score Message (0-49%)'
              }),
              PropertyPaneTextField('errorMessage', {
                label: 'Error Message'
              }),
              PropertyPaneTextField('resultsSavedMessage', {
                label: 'Results Saved Message'
              })
            ]
          },
          {
            groupName: 'Question Management',
            groupFields: [
              {
                key: 'questionManager',
                type: 0,
                targetProperty: 'questionManager',
                properties: {},
                onRender: (): HTMLElement | null => {
                  try {
                    this.disposeReactComponents();
                    this._propertyPaneContainer = document.createElement('div');
                    this._propertyPaneContainer.className = 'quiz-property-pane-container';
                    
                    // Add a setTimeout to ensure the container is properly attached to DOM
                    setTimeout(() => {
                      this.renderPropertyPaneContent();
                    }, 100);
                    
                    return this._propertyPaneContainer;
                  } catch (error) {
                    console.error('Error in onRender:', error);
                    const errorContainer = document.createElement('div');
                    errorContainer.innerHTML = '<div style="color: red; padding: 10px;">Error loading question manager. Please refresh the page and try again.</div>';
                    return errorContainer;
                  }
                },
                onDispose: (): void => {
                  this.disposeReactComponents();
                }
              } as unknown as IPropertyPaneField<IPropertyPaneCustomFieldProps>
            ]
          }
        ]
      }
    ]
  };
}
}