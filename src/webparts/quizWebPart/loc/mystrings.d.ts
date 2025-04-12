declare interface IQuizWebPartStrings {
  // Property pane settings
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  ListNameFieldLabel: string;
  ResultsListNameFieldLabel: string;
  QuestionsPerPageFieldLabel: string;
  DescriptionFieldLabel: string;

  // Environment strings
  AppLocalEnvironmentOffice: string;
  AppOfficeEnvironment: string;
  AppLocalEnvironmentOutlook: string;
  AppOutlookEnvironment: string;
  AppLocalEnvironmentTeams: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  
  // Quiz property pane
  QuizQuestionsTitle: string;
  AddQuestionButtonText: string;
  ImportButtonText: string;
  ShuffleQuestionsButtonText: string;
  FilterQuestionsButtonText: string;
  SortQuestionsButtonText: string;
  DeleteAllQuestionsButtonText: string;
  NoQuestionsMessage: string;
  QuestionAddedSuccess: string;
  QuestionUpdatedSuccess: string;
  QuestionsImportedSuccess: string;
  AllQuestionsDeletedSuccess: string;
  QuestionDeletedSuccess: string;
  QuestionsRandomizedSuccess: string;

  // Add/Edit Question dialog
  AddNewQuestionTitle: string;
  EditQuestionTitle: string;
  QuestionTabText: string;
  AdditionalInfoTabText: string;
  QuestionTitleLabel: string;
  QuestionTitlePlaceholder: string;
  QuestionDescriptionLabel: string;
  QuestionDescriptionPlaceholder: string;
  QuestionTypeLabel: string;
  CategoryLabel: string;
  CategoryPlaceholder: string;
  NewCategoryLabel: string;
  NewCategoryPlaceholder: string;
  ChoicesLabel: string;
  ChoiceMultiSelectLabel: string;
  CorrectAnswerTFLabel: string;
  ChoicePrefix: string;
  CorrectAnswerLabel: string;
  CorrectAnswerPlaceholder: string;
  CaseSensitiveLabel: string;
  PointsLabel: string;
  ExplanationLabel: string;
  ExplanationPlaceholder: string;
  SaveQuestionButtonText: string;
  UpdateQuestionButtonText: string;
  PreviewButtonText: string;
  ResetButtonText: string;
  CancelButtonText: string;
  
  // Validation messages
  ValidationQuestionTitleRequired: string;
  ValidationCategoryRequired: string;
  ValidationMinChoicesRequired: string;
  ValidationCorrectChoiceRequired: string;
  ValidationShortAnswerRequired: string;
  ValidationPointsPositive: string;
  
  // Question Preview
  QuestionPreviewTitle: string;
  PreviewNoAnswer: string;
  PreviewCorrectAnswerLabel: string;
  PreviewCaseSensitiveInfo: string;
  PreviewExplanationLabel: string;
  PreviewNotAvailable: string;
  PreviewCloseButton: string;
  
  // Import Questions dialog
  ImportQuestionsTitle: string;
  ImportQuestionsDescription: string;
  ImportFormatLabel: string;
  CSVFormatOption: string;
  JSONFormatOption: string;
  DefaultCategoryLabel: string;
  NoneCategoryOption: string;
  NewCategoryOption: string;
  NewCategoryNameLabel: string;
  UploadFileLabel: string;
  UseTemplateButton: string;
  CSVFormatGuidelinesTitle: string;
  CSVFormatGuidelinesText: string;
  JSONFormatGuidelinesTitle: string;
  JSONFormatGuidelinesText: string;
  CSVContentLabel: string;
  JSONContentLabel: string;
  CSVContentPlaceholder: string;
  JSONContentPlaceholder: string;
  ImportQuestionsButtonText: string;
  ProcessingLabel: string;
  ImportErrorMessage: string;
  NoQuestionsFoundError: string;
  ImportContentRequiredError: string;
  
  // Filter and sort
  FilterQuestionsTitle: string;
  SortQuestionsTitle: string;
  SelectFilterPlaceholder: string;
  AllQuestionsFilter: string;
  SelectSortMethodPlaceholder: string;
  TitleAscSortOption: string;
  TitleDescSortOption: string;
  CategoryAscSortOption: string;
  CategoryDescSortOption: string;
  ApplyButtonText: string;
  
  // Bulk delete
  ConfirmBulkDeleteTitle: string;
  ConfirmBulkDeleteText: string;
  DeleteButtonText: string;
  
  // Quiz Results
  QuizResultsTitle: string;
  SubmittingLabel: string;
  SuccessDefaultMessage: string;
  ScoreLabel: string;
  ScoreDetailsTemplate: string;
  ExcellentScoreDefaultMessage: string;
  GoodScoreDefaultMessage: string;
  AverageScoreDefaultMessage: string;
  PoorScoreDefaultMessage: string;
  SummaryTabText: string;
  DetailedResultsTabText: string;
  RetakeQuizButtonText: string;
  ViewDetailedResultsButtonText: string;
  
  // Detailed Results
  ResultsNotAvailable: string;
  ScoreFormatTemplate: string;
  ExcellentPerformanceMessage: string;
  GreatPerformanceMessage: string;
  GoodPerformanceMessage: string;
  AveragePerformanceMessage: string;
  PoorPerformanceMessage: string;
  QuestionResultsLabel: string;
  QuestionDetailsLabel: string;
  CorrectResultText: string;
  IncorrectResultText: string;
  YourAnswerLabel: string;
  
  // Question Management
  AddQuestionButtonTextShort: string;
  RefreshButtonText: string;
  DeleteSelectedText: string;
  EditSelectedText: string;
  EditQuestionTooltip: string;
  PreviewTooltip: string;
  DeleteTooltip: string;
  SearchQuestionsPlaceholder: string;
  FilterByCategoryPlaceholder: string;
  FilterByTypePlaceholder: string;
  AllCategoriesOption: string;
  AllTypesOption: string;
  ItemsSelectedMessage: string;
  NoQuestionsFoundMessage: string;
  ConfirmDeleteTitle: string;
  ConfirmDeleteSingleQuestion: string;
  ConfirmDeleteSelectedQuestions: string;
  DeletingLabel: string;
  
  // Quiz Taking UX
  RequireAllQuestionsLabel: string;
  NoQuestionsInCategoryMessage: string;
  SubmittingQuizLabel: string;
  SubmitQuizButtonText: string;
  AllQuestionsRequiredMessage: string;
  
  // Question types display names
  MultipleChoiceType: string;
  TrueFalseType: string;
  MultiSelectType: string;
  ShortAnswerType: string;
  MatchingType: string;
}

declare module 'QuizWebPartStrings' {
  const strings: IQuizWebPartStrings;
  export = strings;
}