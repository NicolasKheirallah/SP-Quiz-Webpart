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
}

export interface IQuizQuestion {
  id: number;
  title: string;
  description?: string;  // New field for detailed question text
  category: string;
  type: QuestionType;
  choices: IChoice[];
  selectedChoice?: string | string[]; // Can be multiple for MultiSelect
  matchingPairs?: { id: string, left: string, right: string }[]; // For matching questions
  correctAnswer?: string; // For short answer
  points?: number; // Optional weighting
  explanation?: string; // Explanation for the answer
  lastModified?: string; // ISO date string for tracking changes
  caseSensitive?: boolean; // For short answer, whether it's case sensitive
  userAnswer?: string | string[];  // Added for tracking user's answer for detailed results
  isCorrect?: boolean;  // Added for tracking if the question was answered correctly
}

export interface IQuizState {
  questions: IQuizQuestion[];
  originalQuestions: IQuizQuestion[]; // For randomization we keep original order
  categories: string[];
  loading: boolean;
  currentPage: number;
  currentCategory: string;
  isAdmin: boolean;
  showResults: boolean;
  score: number;
  totalQuestions: number;
  totalPoints: number;
  answeredQuestions: number; // New field to track progress
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
  // New fields for detailed results
  detailedResults?: IDetailedQuizResults;
  // New field for progress tracking
  quizProgress?: IQuizProgress;
}

// Interface for detailed quiz results
export interface IDetailedQuizResults {
  score: number;
  totalPoints: number;
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
    excellent: string;
    good: string;
    average: string;
    poor: string;
    success: string;
  };
  // New field for detailed results
  detailedResults?: IDetailedQuizResults;
}

export interface IQuizQuestionProps {
  question: IQuizQuestion;
  onAnswerSelect: (questionId: number, choiceId: string | string[]) => void;
  questionNumber: number;
  totalQuestions: number;
  showProgressIndicator: boolean;
}

export interface IAddQuestionFormProps {
  categories: string[];
  onSubmit: (newQuestion: IQuizQuestion) => void;
  onCancel: () => void;
  isSubmitting: boolean;
  onPreviewQuestion: (question: IQuizQuestion) => void;
  initialQuestion?: IQuizQuestion; // Added for editing existing questions
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