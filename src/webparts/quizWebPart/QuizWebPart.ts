import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneField
} from '@microsoft/sp-property-pane';
import { IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'QuizWebPartStrings';
import Quiz from './components/Quiz';
import { IQuizProps } from './components/IQuizProps';
import { IQuizQuestion, QuestionType } from './components/interfaces';
import QuizPropertyPane from './components/QuizPropertyPane';

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
}

export default class QuizWebPart extends BaseClientSideWebPart<IQuizWebPartProps> {
  private _reactElement: HTMLElement | null = null;
  private _propertyPaneContainer: HTMLElement | null = null;

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
        questions: this.properties.questions || [],
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
      // Ensure proper container exists
      if (!this._propertyPaneContainer) {
        this._propertyPaneContainer = document.createElement('div');
        this._propertyPaneContainer.className = 'quiz-property-pane-container';
      }

      const propertyPaneElement = React.createElement(
        QuizPropertyPane,
        {
          questions: this.properties.questions || [],
          onUpdateQuestions: (questions: IQuizQuestion[]) => {
            this.properties.questions = questions;
            this.render();
          }
        }
      );

      // Render the new property pane component
      ReactDom.render(propertyPaneElement, this._propertyPaneContainer);
    }
  }

  protected onDispose(): void {
    // Ensure we unmount all React components to prevent memory leaks
    this.disposeReactComponents();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
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

    // Clean up any existing components
    this.disposeReactComponents();

    // Create the container for the property pane component
    this._propertyPaneContainer = document.createElement('div');
    this._propertyPaneContainer.className = 'quiz-property-pane-container';
  }

  protected onPropertyPaneConfigurationEnd(): void {
    // Clean up any property pane components
    this.disposeReactComponents();
  }

  // Handle property pane changes
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    // If questions were changed, trigger re-render
    if (propertyPath === 'questions') {
      this.render();
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

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
                  label: 'Time Limit (minutes)',
                  description: 'Maximum time allowed for the quiz (0 for no time limit)'
                })
              ]
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
                    this.disposeReactComponents();
                    this._propertyPaneContainer = document.createElement('div');
                    this._propertyPaneContainer.className = 'quiz-property-pane-container';
                    
                    return this._propertyPaneContainer;
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