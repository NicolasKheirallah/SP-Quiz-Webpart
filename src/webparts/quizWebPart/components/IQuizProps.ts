import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IQuizQuestion } from './interfaces';

export interface IQuizProps {
  title: string;
  questionsPerPage: number;
  context: WebPartContext;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
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
  passingScore?: number;
  timeLimit?: number;
  enableQuestionTimeLimit: boolean;
  defaultQuestionTimeLimit: number;
  questions: IQuizQuestion[];
  resultsListName: string; // Added property for the list name
  updateQuestions: (questions: IQuizQuestion[]) => void;
}