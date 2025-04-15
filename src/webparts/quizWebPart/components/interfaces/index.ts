import { WebPartContext } from "@microsoft/sp-webpart-base";

export enum QuestionType {
  MultipleChoice = 'multipleChoice',
  TrueFalse = 'trueFalse',
  MultiSelect = 'multiSelect',
  ShortAnswer = 'shortAnswer',
  Matching = 'matching'
}

export interface IChoice {
  id: string;
  text: string;
  isCorrect: boolean;
  image?: IQuizImage; // Optional image for this choice
}

// Image interface
export interface IQuizImage {
  id: string;
  url: string;
  altText?: string;
  fileName: string;
  width?: number;
  height?: number;
}

// Code snippet interface
export interface ICodeSnippet {
  id: string;
  code: string;
  language: string; // e.g., 'javascript', 'python', 'csharp', etc.
  lineNumbers?: boolean;
  highlightLines?: number[]; // Lines to highlight
}

export interface IQuizQuestion {
  id: number;
  title: string;
  description?: string;  // New field for detailed question text
  category: string;
  type: QuestionType;
  choices: IChoice[];
  selectedChoice?: string | string[]; // Can be multiple for MultiSelect
  matchingPairs?: IMatchingPair[];
  correctAnswer?: string; // For short answer
  points?: number; // Optional weighting
  explanation?: string; // Explanation for the answer
  lastModified?: string; // ISO date string for tracking changes
  caseSensitive?: boolean; // For short answer, whether it's case sensitive
  userAnswer?: string | string[];  // Added for tracking user's answer for detailed results
  isCorrect?: boolean;  // Added for tracking if the question was answered correctly
  images?: IQuizImage[];
  timeLimit?: number; // Time limit in seconds for this specific question
  codeSnippets?: ICodeSnippet[]; // For code syntax highlighting
}

export interface IQuizState {
  questions: IQuizQuestion[];
  originalQuestions: IQuizQuestion[];
  categories: string[];
  loading: boolean;
  currentPage: number;
  currentCategory: string;
  isAdmin: boolean;
  showResults: boolean;
  score: number;
  totalQuestions: number;
  totalPoints: number;
  answeredQuestions: number; 
  isSubmitting: boolean;
  submissionSuccess: boolean;
  submissionError: string;
  showAddQuestionForm: boolean;
  newQuestion: IQuizQuestion;
  previewQuestion: IQuizQuestion | undefined;
  showQuestionPreview: boolean;
  importDialogOpen: boolean;
  exportDialogOpen?: boolean;
  showConfirmDialog?: boolean;
  confirmDialogAction?: string;
  adminView?: string;
  submitRequireAllAnswered?: boolean;
  showEditQuestionsDialog: boolean;
  detailedResults?: IDetailedQuizResults;
  quizProgress?: IQuizProgress;
  showStartPage: boolean;
  quizStarted: boolean;
  overallTimerExpired: boolean;
  expiredQuestions: number[];
  showSaveProgressDialog: boolean;
  hasSavedProgress: boolean;
  showResumeDialog: boolean;
  savedProgressId?: number;
  timeRemaining?: number;
  correctlyAnsweredQuestions?: number;
  showCategoryOrderDialog: boolean;

}


// Interface for detailed quiz results
export interface IDetailedQuizResults {
  score: number;
  totalPoints: number;
  totalQuestions: number;
  answeredQuestions: number;
  correctlyAnsweredQuestions: number; 
  percentageAnswered: number;
  percentageCorrect: number;
  percentageCorrectOfAnswered: number;
  percentage: number;
  questionResults: IQuestionResult[];
  timestamp: string;
}



export interface IQuestionResult {
  id: number;
  title: string;
  userAnswer: string | string[] | undefined;
  correctAnswer: string | string[] | undefined;
  isCorrect: boolean;
  points: number;
  earnedPoints: number;
  explanation?: string;
}

// Interface for progress tracking
export interface IQuizProgress {
  currentQuestion: number;
  totalQuestions: number;
  answeredQuestions: number;
  percentage: number;
  remainingTime?: number;  // For timed quizzes
}

export interface IQuizResultsProps {
  score: number;
  totalQuestions: number;
  totalPoints: number;
  isSubmitting: boolean;
  submissionSuccess: boolean;
  submissionError: string;
  onRetakeQuiz: () => void;
  messages: {
    excellent?: string;
    good?: string;
    average?: string;
    poor?: string;
    success?: string;
  };
  detailedResults?: IDetailedQuizResults;
}


export interface IQuizQuestionProps {
  question: IQuizQuestion;
  onAnswerSelect: (questionId: number, choiceId: string | string[]) => void;
  questionNumber: number;
  totalQuestions: number;
  showProgressIndicator?: boolean;
  onTimeExpired?: (questionId: number) => void;
}

export interface IAddQuestionFormProps {
  categories: string[];
  onSubmit: (newQuestion: IQuizQuestion) => void;
  onCancel: () => void;
  isSubmitting: boolean;
  onPreviewQuestion: (question: IQuizQuestion) => void;
  initialQuestion?: IQuizQuestion; // Added for editing existing questions
  context?: WebPartContext; // SPFx context for file picker
  defaultQuestionTimeLimit?: number; // Add this property
}

export interface IImportQuestionsProps {
  onImportQuestions: (questions: IQuizQuestion[]) => void;
  onCancel: () => void;
  existingCategories: string[];
}

export interface IQuestionPreviewProps {
  question: IQuizQuestion;
  onClose: () => void;
}

export interface IQuizProgressTrackerProps {
  progress: IQuizProgress;
  showPercentage?: boolean;
  showNumbers?: boolean;
  showIcon?: boolean;
  showTimer?: boolean;
}

// Interface for image upload component
export interface IImageUploadProps {
  onImageUpload: (image: IQuizImage) => void;
  onImageRemove?: () => void;
  currentImage?: IQuizImage;
  label?: string;
  accept?: string;
  maxSizeMB?: number;
  context?: WebPartContext;
}

// Interface for question timer component
export interface IQuestionTimerProps {
  timeLimit: number; // in seconds
  onTimeExpired: () => void;
  paused?: boolean;
  warningThreshold?: number; // percentage when to start warning (default: 20%)
  criticalThreshold?: number; // percentage when to show critical warning (default: 10%)
  showText?: boolean;
}

// Interface for code snippet component
export interface ICodeSnippetProps {
  snippet?: ICodeSnippet;
  onChange: (snippet: ICodeSnippet) => void;
  onRemove?: () => void;
  isEditing?: boolean;
  label?: string;
}
export interface IQuizPropertyPaneProps {
  questions: IQuizQuestion[];
  onUpdateQuestions: (questions: IQuizQuestion[]) => void;
  context: WebPartContext;
}

export interface IMatchingPair {
  id: string;
  leftItem: string;
  rightItem: string;
  userSelectedRightId?: string;
}

export interface ISavedQuizProgress {
  id?: number;
  userId: string;
  userName: string;
  quizTitle: string;
  questions: IQuizQuestion[];
  lastSaved: string;
  timeRemaining?: number;
  currentPage: number;
  currentCategory: string;
}

export interface IQuizStartPageProps {
  title: string;
  onStartQuiz: () => void;
  totalQuestions: number;
  totalPoints: number;
  categories: string[];
  timeLimit?: number; // in seconds
  passingScore?: number; // percentage
  quizImage?: string; // optional image URL
  description?: string;
  hasSavedProgress?: boolean;
  onResumeQuiz?: () => void;

}
