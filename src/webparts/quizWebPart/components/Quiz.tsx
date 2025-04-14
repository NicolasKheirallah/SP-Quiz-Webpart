import * as React from 'react';
import { IQuizProps } from './IQuizProps';
import { IQuizState, IQuizQuestion, QuestionType, ISavedQuizProgress } from './interfaces';
import { IDetailedQuizResults, IQuestionResult } from './interfaces';
import QuizProgressTracker from './QuizProgressTracker';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { v4 as uuidv4 } from 'uuid';
import QuizQuestion from './QuizQuestion';
import QuizResults from './QuizResults';
import AddQuestionDialog from './AddQuestionDialog';
import ImportQuestionsDialog from './ImportQuestionsDialog';
import QuestionPreview from './QuestionPreview';
import { DisplayMode } from '@microsoft/sp-core-library';
import QuizStartPage from './QuizStartPage';
import QuizTimer from './QuizTimer';
import {
  Spinner,
  SpinnerSize,
  Stack,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Text,
  PrimaryButton,
  DefaultButton,
  Pivot,
  PivotItem,
  Dialog,
  DialogType,
  DialogFooter,
  IIconProps,
  IStackStyles,
  Checkbox,
} from '@fluentui/react';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import styles from './Quiz.module.scss';
import QuestionManagement from './QuestionManagement';
import { SPHttpClient } from '@microsoft/sp-http';

// Icons
const editIcon: IIconProps = { iconName: 'Edit' };
const addIcon: IIconProps = { iconName: 'Add' };
const importIcon: IIconProps = { iconName: 'Download' };
const submitIcon: IIconProps = { iconName: 'CheckMark' };
const exportIcon: IIconProps = { iconName: 'Upload' };
const saveIcon: IIconProps = { iconName: 'Save' };

// Styles
const mainContainerStyles: IStackStyles = {
  root: {
    padding: '20px',
    maxWidth: '1200px',
    margin: '0 auto'
  }
};

const stackTokens: IStackTokens = {
  childrenGap: 12
};

export default class Quiz extends React.Component<IQuizProps, IQuizState> {
  constructor(props: IQuizProps) {
    super(props);
    this.handleAddQuestion = this.handleAddQuestion.bind(this);
    this.handleEditQuestion = this.handleEditQuestion.bind(this);
    this.handleDeleteQuestion = this.handleDeleteQuestion.bind(this);
    this.handleImportQuestions = this.handleImportQuestions.bind(this);
    this.handleExportQuestions = this.handleExportQuestions.bind(this);
    this.handlePreviewQuestion = this.handlePreviewQuestion.bind(this);
    this.handleOpenEditQuestionsDialog = this.handleOpenEditQuestionsDialog.bind(this);
    this.handleCloseEditQuestionsDialog = this.handleCloseEditQuestionsDialog.bind(this);

    // Get unique categories from questions
    const categoriesSet = new Set<string>();
    props.questions.forEach(q => {
      if (q.category) categoriesSet.add(q.category);
    });
    const categories = ['All', ...Array.from(categoriesSet)];

    this.state = {
      questions: [...props.questions], // Clone to avoid direct mutation
      originalQuestions: [...props.questions], // Store original order
      categories,
      loading: false,
      currentPage: 1,
      currentCategory: 'All',
      isAdmin: false, // This will be removed/ignored, using displayMode instead
      showResults: false,
      score: 0,
      totalQuestions: 0,
      totalPoints: 0,
      answeredQuestions: 0,
      isSubmitting: false,
      submissionSuccess: false,
      submissionError: '',
      showAddQuestionForm: false,
      newQuestion: this.getEmptyQuestion(),
      previewQuestion: undefined,
      showQuestionPreview: false,
      importDialogOpen: false,
      exportDialogOpen: false,
      showConfirmDialog: false,
      confirmDialogAction: '',
      adminView: 'questions',
      submitRequireAllAnswered: false,
      showEditQuestionsDialog: false,
      showStartPage: true,
      quizStarted: false,
      overallTimerExpired: false,
      expiredQuestions: [],
      showSaveProgressDialog: false,
      hasSavedProgress: false,
      showResumeDialog: false



    };
  }

  private getEmptyQuestion(): IQuizQuestion {
    return {
      id: Date.now(),
      title: '',
      category: '',
      type: QuestionType.MultipleChoice,
      choices: [
        { id: uuidv4(), text: '', isCorrect: false },
        { id: uuidv4(), text: '', isCorrect: false },
      ],
      matchingPairs: []
    };
  }


  public componentDidMount(): void {
    this.randomizeQuestionsIfNeeded();
    this.checkForSavedProgress().catch((error) =>
      console.error("Error in checkForSavedProgress:", error)
    );
  }

  public componentDidUpdate(prevProps: IQuizProps): void {
    // If randomize setting changes, update questions
    if (prevProps.randomizeQuestions !== this.props.randomizeQuestions ||
      prevProps.randomizeAnswers !== this.props.randomizeAnswers) {
      this.randomizeQuestionsIfNeeded();
    }

    // If questions array from props changes
    if (prevProps.questions !== this.props.questions) {
      // Get unique categories
      const categoriesSet = new Set<string>();
      this.props.questions.forEach(q => {
        if (q.category) categoriesSet.add(q.category);
      });
      const categories = ['All', ...Array.from(categoriesSet)];

      this.setState({
        questions: [...this.props.questions],
        originalQuestions: [...this.props.questions],
        categories
      }, () => this.randomizeQuestionsIfNeeded());
    }
  }
  private handleOpenEditQuestionsDialog = (): void => {
    this.setState({ showEditQuestionsDialog: true });
  }

  private handleCloseEditQuestionsDialog = (): void => {
    this.setState({ showEditQuestionsDialog: false });
  }

  private randomizeQuestionsIfNeeded(): void {
    const { randomizeQuestions, randomizeAnswers } = this.props;
    let updatedQuestions = [...this.state.originalQuestions];

    if (randomizeQuestions) {
      // Randomize question order
      updatedQuestions = this.shuffleArray([...updatedQuestions]);
    }

    if (randomizeAnswers) {
      // Randomize answer choices for each question
      updatedQuestions = updatedQuestions.map(question => {
        const shuffledChoices = this.shuffleArray([...question.choices]);
        return { ...question, choices: shuffledChoices };
      });
    }

    this.setState({ questions: updatedQuestions });
  }
  private getTimeRemaining(): number {
    return this.state.timeRemaining || 0;

  }

  private setRemainingTime(seconds: number): void {
    this.setState({ timeRemaining: seconds });
  }

  private handleSaveProgressClick = (): void => {
    this.setState({ showSaveProgressDialog: true });
  };

  private handleSaveProgressCancel = (): void => {
    this.setState({ showSaveProgressDialog: false });
  };

  private handleSaveProgressConfirm = async (): Promise<void> => {
    try {
      await this.saveQuizProgress();
    } catch (error) {
      console.error("Error saving quiz progress:", error);
      this.setState({
        submissionError: error instanceof Error ? error.message : 'Failed to save progress',
        isSubmitting: false
      });
    }
  };



  private discardSavedProgress = async (): Promise<void> => {
    try {
      if (!this.state.savedProgressId) return;

      const { context } = this.props;
      const spHttpClient = context.spHttpClient;
      const webUrl = context.pageContext.web.absoluteUrl;
      const progressListName = "QuizProgress";

      // Delete the saved progress
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${this.state.savedProgressId})`;

      await spHttpClient.fetch(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          method: 'DELETE',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      this.setState({
        savedProgressId: undefined,
        hasSavedProgress: false,
        showResumeDialog: false
      });

      // Start a new quiz
      this.handleStartQuiz();

    } catch (error) {
      console.error('Error discarding saved progress:', error);
      // Handle the error or just start a new quiz anyway
      this.setState({ showResumeDialog: false });
      this.handleStartQuiz();
    }
  };


  // This is the updated saveQuizResults method to use the selected list name
  private saveQuizResults = async (): Promise<boolean> => {
    try {
      const spHttpClient = this.props.context.spHttpClient;
      const webUrl = this.props.context.pageContext.web.absoluteUrl;
      const currentUser = this.props.context.pageContext.user;

      // Use the list name from props (selected in property pane)
      const resultsListName = this.props.resultsListName || 'QuizResults';

      // Calculate the score percentage properly
      const scorePercentage = this.state.totalPoints > 0
        ? Math.round((this.state.score / this.state.totalPoints) * 100)
        : 0;

      // Prepare the question results data with proper scoring information
      const questionResults = this.state.questions.map(question => {
        // Determine if the question was answered correctly
        const isCorrect = this.isQuestionCorrect(question);
        const points = question.points || 1;
        const earnedPoints = isCorrect ? points : 0;

        return {
          QuestionId: question.id.toString(),
          QuestionTitle: question.title,
          QuestionType: question.type,
          SelectedChoice: question.selectedChoice
            ? (Array.isArray(question.selectedChoice)
              ? question.selectedChoice.join(',')
              : question.selectedChoice.toString())
            : '',
          IsCorrect: isCorrect,
          EarnedPoints: earnedPoints,
          PossiblePoints: points
        };
      });

      // Prepare detailed result data
      const resultData = {
        Title: `Quiz Result - ${new Date().toLocaleDateString()}`,
        UserName: currentUser.displayName || 'Anonymous',
        UserEmail: currentUser.email || 'Not provided',
        UserId: currentUser.loginName || 'Unknown',
        QuizTitle: this.props.title || 'SharePoint Quiz',

        // Score details - ensuring we have valid numbers
        Score: this.state.score,
        TotalPoints: this.state.totalPoints,
        ScorePercentage: scorePercentage,
        QuestionsAnswered: this.state.answeredQuestions || 0,
        TotalQuestions: this.state.questions.length || 0,

        // Timestamp
        ResultDate: new Date().toISOString(),

        // Detailed question results
        QuestionDetails: JSON.stringify(questionResults)
      };

      console.log(`Saving quiz results to list: ${resultsListName}`, resultData);

      // Save result to the selected list
      try {
        const response = await spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${resultsListName}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            },
            body: JSON.stringify(resultData)
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Error response from SharePoint:', errorText);
          throw new Error(`Failed to save quiz results: ${errorText}`);
        }

        const responseData = await response.json();
        console.log(`Quiz results saved successfully to ${resultsListName}:`, responseData);

        return true;
      } catch (error) {
        console.error(`Error saving to ${resultsListName} list:`, error);
        throw error;
      }
    } catch (error) {
      console.error('Error in saveQuizResults:', error);

      // Update state with submission error
      this.setState({
        submissionError: error instanceof Error
          ? error.message
          : 'An unexpected error occurred while saving results'
      });

      return false;
    }
  }


  private getCorrectAnswerText = (question: IQuizQuestion): string | string[] | undefined => {
    switch (question.type) {
      case QuestionType.MultipleChoice:
      case QuestionType.TrueFalse: {
        const correctChoice = question.choices.find(c => c.isCorrect);
        return correctChoice ? correctChoice.text : undefined;
      }
      case QuestionType.MultiSelect: {
        const correctChoices = question.choices.filter(c => c.isCorrect).map(c => c.text);
        return correctChoices.length > 0 ? correctChoices : undefined;
      }
      case QuestionType.ShortAnswer:
        return question.correctAnswer;
      case QuestionType.Matching: {
        if (question.matchingPairs) {
          return question.matchingPairs.map(pair => `${pair.leftItem} → ${pair.rightItem}`);
        }
        return undefined;
      }
      default:
        return undefined;
    }
  };


  private checkForSavedProgress = async (): Promise<void> => {
    try {
      // Only check for saved progress in read mode
      if (this.props.displayMode !== DisplayMode.Read) return;

      console.log("Checking for saved progress...");
      const { context } = this.props;
      const spHttpClient = context.spHttpClient;
      const webUrl = context.pageContext.web.absoluteUrl;
      const currentUser = context.pageContext.user;

      // Use the same list as quiz results or a dedicated progress list
      const progressListName = "QuizProgress";

      // Query for saved progress for this user and quiz
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items?$filter=UserId eq '${currentUser.loginName}' and QuizTitle eq '${this.props.title}'&$orderby=Modified desc&$top=1`;

      console.log(`Checking saved progress at endpoint: ${endpoint}`);

      try {
        const response = await spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }
        );

        if (!response.ok) {
          const errorText = await response.text();
          console.error(`Error response when checking for saved progress: ${errorText}`);
          return;
        }

        const data = await response.json();
        console.log("Saved progress check response:", data);

        if (data.value && data.value.length > 0) {
          // Found saved progress
          const savedProgress = data.value[0];
          console.log("Found saved progress:", savedProgress);

          this.setState({
            hasSavedProgress: true,
            savedProgressId: savedProgress.Id
          });

          // Only show resume dialog in read mode and on start page
          if (this.state.showStartPage) {
            this.setState({ showResumeDialog: true });
          }
        } else {
          console.log("No saved progress found");
        }
      } catch (error) {
        console.error('Error in API call when checking for saved progress:', error);
      }
    } catch (error) {
      console.error('Error checking for saved progress:', error);
    }
  };

  private saveQuizProgress = async (): Promise<boolean> => {
    try {
      this.setState({ isSubmitting: true });

      const { context } = this.props;
      const spHttpClient = context.spHttpClient;
      const webUrl = context.pageContext.web.absoluteUrl;
      const currentUser = context.pageContext.user;

      // Use a dedicated list for progress
      const progressListName = "QuizProgress";

      // Prepare progress data
      const progressData: ISavedQuizProgress = {
        userId: currentUser.loginName || 'Unknown',
        userName: currentUser.displayName || 'Anonymous',
        quizTitle: this.props.title || 'SharePoint Quiz',
        questions: this.state.questions,
        lastSaved: new Date().toISOString(),
        timeRemaining: this.props.timeLimit ? this.getTimeRemaining() : undefined,
        currentPage: this.state.currentPage,
        currentCategory: this.state.currentCategory
      };
      // Convert to SharePoint item format
      const spItemData = {
        Title: `${this.props.title} - ${currentUser.displayName} - In Progress`,
        UserId: progressData.userId,
        UserName: progressData.userName,
        QuizTitle: progressData.quizTitle,
        QuizData: JSON.stringify(progressData),
        LastSaved: progressData.lastSaved
      };

      console.log("Saving quiz progress:", {
        listName: progressListName,
        progressData: progressData,
        spItemData: spItemData,
        savedProgressId: this.state.savedProgressId
      });

      // Decide whether to create or update based on savedProgressId
      let endpoint = '';
      let method = '';
      const headers: HeadersInit = {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      };

      if (this.state.savedProgressId) {
        // Update existing record
        endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${this.state.savedProgressId})`;
        method = 'POST';  // Use POST instead of MERGE
        headers['X-HTTP-Method'] = 'MERGE';  // Add X-HTTP-Method override
        headers['IF-MATCH'] = '*';
      } else {
        // Create new record
        endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items`;
        method = 'POST';
      }

      console.log(`Saving to endpoint: ${endpoint} with method: ${method}`);

      try {
        const response = await spHttpClient.fetch(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            method,
            headers,
            body: JSON.stringify(spItemData)
          }
        );

        // Log response status
        console.log(`Response status: ${response.status}`);

        if (!response.ok) {
          const errorData = await response.json();
          console.error('Error response from SharePoint:', errorData);
          throw new Error(`Failed to save progress: ${JSON.stringify(errorData)}`);
        }

        // If this is a new record, get the ID
        if (!this.state.savedProgressId) {
          const responseData = await response.json();
          console.log("Save response data:", responseData);
          this.setState({ savedProgressId: responseData.Id });
        }

        // Display success message
        this.setState({
          hasSavedProgress: true,
          isSubmitting: false,
          showSaveProgressDialog: false,
          submissionSuccess: true
        });

        // Show temporary success message
        alert("Your progress has been saved successfully! You can resume this quiz later.");

        return true;
      } catch (error) {
        console.error('Error in API call:', error);
        throw error; // Re-throw to be caught by outer catch
      }
    } catch (error) {
      console.error('Error saving quiz progress:', error);

      // Create a user-friendly error message
      let errorMessage = 'Error saving progress';
      if (error instanceof Error) {
        errorMessage = `Error: ${error.message}`;
      }

      this.setState({
        submissionError: errorMessage,
        isSubmitting: false,
        showSaveProgressDialog: false
      });

      // Show error to user
      alert(`Failed to save quiz progress: ${errorMessage}`);

      return false;
    }
  };

  /**
   * Resumes a previously saved quiz session from SharePoint
   * Retrieves the saved state and restores it to the component
   */
  private resumeQuiz = async (): Promise<void> => {
    try {
      this.setState({
        loading: true,
        submissionError: '',
        showResumeDialog: false // Immediately close the resume dialog
      });

      const { context } = this.props;
      const spHttpClient = context.spHttpClient;
      const webUrl = context.pageContext.web.absoluteUrl;

      // Get saved progress from the QuizProgress list
      const progressListName = "QuizProgress";
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${this.state.savedProgressId})`;

      console.log(`Fetching saved quiz progress from: ${endpoint}`);

      const response = await spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Error response when retrieving saved progress: ${errorText}`);
        throw new Error(`Failed to retrieve saved progress: ${response.status} ${response.statusText}`);
      }

      const savedItem = await response.json();
      console.log("Retrieved saved quiz item:", savedItem);

      if (!savedItem.QuizData) {
        throw new Error("Saved progress data is missing or corrupted");
      }

      // Parse the saved quiz data
      let progressData: ISavedQuizProgress;
      try {
        progressData = JSON.parse(savedItem.QuizData) as ISavedQuizProgress;
        console.log("Parsed quiz progress data:", progressData);
      } catch (parseError) {
        console.error("Error parsing saved QuizData JSON:", parseError);
        throw new Error("Could not parse saved quiz data. The data may be corrupted.");
      }

      // Validate and merge saved questions with current question set
      const mergedQuestions = progressData.questions
        .filter(savedQuestion => this.isValidQuizQuestion(savedQuestion))
        .map((savedQuestion: Partial<IQuizQuestion>) => {
          // Find the matching original question to preserve any updated properties
          const originalQuestion = this.state.originalQuestions.find(q => q.id === savedQuestion.id);

          if (originalQuestion) {
            // Combine the original question with the saved one, prioritizing saved data
            return {
              ...originalQuestion,
              ...savedQuestion,
              // Ensure selected choices are preserved
              selectedChoice: savedQuestion.selectedChoice
            } as IQuizQuestion;
          } else {
            // If we can't find the original question, just use the saved question
            return savedQuestion as IQuizQuestion;
          }
        });

      if (mergedQuestions.length === 0) {
        throw new Error("No valid questions found in saved quiz data");
      }

      // Count answered questions
      const answeredQuestions = mergedQuestions.filter(q => q.selectedChoice !== undefined).length;

      // Restore quiz state
      this.setState({
        questions: mergedQuestions,
        currentPage: progressData.currentPage || 1,
        currentCategory: progressData.currentCategory || 'All',
        showStartPage: false,
        quizStarted: true,
        loading: false,
        answeredQuestions,
        showResumeDialog: false
      });

      // If there was a time limit, restore remaining time
      if (progressData.timeRemaining !== undefined && this.props.timeLimit) {
        this.setRemainingTime(progressData.timeRemaining);
      }

      console.log("Quiz successfully resumed with", mergedQuestions.length, "questions");

      // Delete the saved progress after successful resume
      await this.deleteSavedProgress(this.state.savedProgressId);

    } catch (error) {
      console.error('Error resuming quiz:', error);

      // Create a user-friendly error message
      let errorMessage = 'Failed to resume quiz';
      if (error instanceof Error) {
        errorMessage = `Failed to resume quiz: ${error.message}`;
      }

      this.setState({
        loading: false,
        submissionError: errorMessage,
        showResumeDialog: false
      });

      // Show error to user
      alert(`${errorMessage} Starting a new quiz instead.`);

      // Fall back to starting a new quiz
      this.handleStartQuiz();
    }
  };

  // Validation method
  private isValidQuizQuestion(obj: unknown): obj is IQuizQuestion {
    if (!obj || typeof obj !== 'object') return false;

    const questionCandidate = obj as Partial<IQuizQuestion>;

    return (
      // Check for required properties
      typeof questionCandidate.id === 'number' &&
      typeof questionCandidate.title === 'string' &&
      typeof questionCandidate.category === 'string' &&
      typeof questionCandidate.type === 'string' &&
      Array.isArray(questionCandidate.choices) &&
      // Additional optional checks
      questionCandidate.choices.every(choice =>
        typeof choice === 'object' &&
        typeof choice.id === 'string' &&
        typeof choice.text === 'string' &&
        typeof choice.isCorrect === 'boolean'
      )
    );
  }


  private deleteSavedProgress = async (progressId?: number): Promise<void> => {
    if (!progressId) return;

    try {
      const { context } = this.props;
      const spHttpClient = context.spHttpClient;
      const webUrl = context.pageContext.web.absoluteUrl;
      const progressListName = "QuizProgress";

      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${progressId})`;

      await spHttpClient.fetch(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          method: 'DELETE',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      console.log(`Saved progress ${progressId} deleted successfully`);

      // Reset saved progress state
      this.setState({
        savedProgressId: undefined,
        hasSavedProgress: false
      });
    } catch (error) {
      console.error('Error deleting saved progress:', error);
    }
  };


  private renderSaveProgressDialog = (): JSX.Element | null => {
    if (!this.state.showSaveProgressDialog) return null;

    return (
      <Dialog
        hidden={false}
        onDismiss={this.handleSaveProgressCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Save Progress',
          subText: 'Do you want to save your progress and continue later? You can resume from where you left off.'
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 450 } }
        }}
      >
        {this.state.submissionError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            dismissButtonAriaLabel="Close"
            styles={{ root: { marginBottom: 15 } }}
          >
            {this.state.submissionError}
          </MessageBar>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={this.handleSaveProgressConfirm}
            text="Save Progress"
            disabled={this.state.isSubmitting}
          />
          <DefaultButton
            onClick={this.handleSaveProgressCancel}
            text="Cancel"
            disabled={this.state.isSubmitting}
          />
          {this.state.isSubmitting && (
            <Spinner size={SpinnerSize.small} style={{ marginLeft: 10 }} />
          )}
        </DialogFooter>
      </Dialog>
    );
  };


  // Add dialog rendering for resume quiz
  private renderResumeDialog = (): JSX.Element | null => {
    if (!this.state.showResumeDialog) return null;

    return (
      <Dialog
        hidden={false}
        onDismiss={() => this.setState({ showResumeDialog: false })}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Resume Quiz',
          subText: 'You have a saved quiz in progress. Would you like to resume where you left off?'
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 450 } }
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={this.resumeQuiz} text="Resume Quiz" />
          <DefaultButton onClick={this.discardSavedProgress} text="Start New Quiz" />
        </DialogFooter>
      </Dialog>
    );
  };

  private handleQuestionTimeExpired = (questionId: number): void => {
    // Add this expired question ID to a tracking array
    this.setState(prevState => ({
      expiredQuestions: [...(prevState.expiredQuestions || []), questionId]
    }));

    // Optionally, you can choose to automatically advance to the next question
    // This is commented out by default, but can be enabled based on requirements
    /*
    const { currentPage, filteredQuestions, questionsPerPage } = this.state;
    const totalPages = Math.ceil(filteredQuestions.length / questionsPerPage);
    
    if (currentPage < totalPages) {
      // Move to next page if not on the last page
      this.handlePageChange(currentPage + 1);
    }
    */

    // You could also choose to mark this question with a default answer
    // For example, for multiple choice questions, you might select the first option
    /*
    const { questions } = this.state;
    const questionIndex = questions.findIndex(q => q.id === questionId);
    
    if (questionIndex !== -1) {
      const question = questions[questionIndex];
      let defaultAnswer: string | string[] | undefined = undefined;
      
      // Determine default answer based on question type
      switch (question.type) {
        case QuestionType.MultipleChoice:
        case QuestionType.TrueFalse:
          // Select first option
          defaultAnswer = question.choices[0]?.id;
          break;
        case QuestionType.MultiSelect:
          // Select no options
          defaultAnswer = [];
          break;
        case QuestionType.ShortAnswer:
          // Empty string
          defaultAnswer = '';
          break;
      }
      
      // If we have a default answer, update the question
      if (defaultAnswer !== undefined) {
        const updatedQuestions = [...questions];
        updatedQuestions[questionIndex].selectedChoice = defaultAnswer;
        this.setState({ questions: updatedQuestions });
      }
    }
    */
  }

  // Helper to determine if a question is correct
  private isQuestionCorrect = (question: IQuizQuestion): boolean => {
    // If the question wasn't answered, it's not correct
    if (question.selectedChoice === undefined) {
      return false;
    }

    try {
      switch (question.type) {
        case QuestionType.MultipleChoice:
        case QuestionType.TrueFalse: {
          const selectedChoiceId = question.selectedChoice as string;
          const correctChoice = question.choices.find(c => c.id === selectedChoiceId && c.isCorrect);
          return !!correctChoice;
        }

        case QuestionType.MultiSelect: {
          // For multi-select, all correct choices must be selected and no incorrect choices
          if (!Array.isArray(question.selectedChoice)) {
            return false; // Invalid data type for multi-select
          }

          const selectedIds = new Set(question.selectedChoice);
          const correctChoices = question.choices.filter(c => c.isCorrect);

          // If there are no correct choices defined, the question is invalid
          if (correctChoices.length === 0) {
            return false;
          }

          // Make sure all correct choices are selected
          const allCorrectChoicesSelected = correctChoices.every(choice =>
            selectedIds.has(choice.id)
          );

          // Make sure no incorrect choices are selected
          const noIncorrectChoicesSelected = question.choices
            .filter(choice => !choice.isCorrect)
            .every(choice => !selectedIds.has(choice.id));

          return allCorrectChoicesSelected && noIncorrectChoicesSelected;
        }

        case QuestionType.ShortAnswer: {
          if (typeof question.selectedChoice !== 'string' ||
            typeof question.correctAnswer !== 'string') {
            return false;
          }

          const userAnswer = (question.selectedChoice as string).trim();
          const correctAnswer = (question.correctAnswer as string).trim();

          // Handle case sensitivity
          return question.caseSensitive === true
            ? userAnswer === correctAnswer
            : userAnswer.toLowerCase() === correctAnswer.toLowerCase();
        }

        // Add this case for matching questions
        case QuestionType.Matching: {
          // For matching questions, we need to verify that all pairs are correctly matched
          if (!Array.isArray(question.selectedChoice) || !question.matchingPairs) {
            return false;
          }

          // The selectedChoice array contains strings in the format "leftId:rightId"
          // Parse these into a map for easier checking
          const userSelections = new Map<string, string>();
          question.selectedChoice.forEach(selection => {
            if (typeof selection === 'string') {
              const [leftId, rightId] = selection.split(':');
              userSelections.set(leftId, rightId);
            }
          });

          return question.matchingPairs.every(pair => {
            const userSelectedRightId = userSelections.get(pair.id);
            return userSelectedRightId === pair.id;
          });
        }

        default:
          return false;
      }
    } catch (error) {
      console.error(`Error checking if question ${question.id} is correct:`, error);
      return false; // Fail safe
    }
  };

  // Fisher-Yates shuffle algorithm
  private shuffleArray<T>(array: T[]): T[] {
    const newArray = [...array];
    for (let i = newArray.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
    }
    return newArray;
  }

  private handleCategoryChange = (item?: PivotItem): void => {
    if (item) {
      this.setState({
        currentCategory: item.props.itemKey || 'All',
        currentPage: 1
      });
    }
  };
  private handleStartQuiz = (): void => {
    this.setState({
      showStartPage: false,
      quizStarted: true,
      currentPage: 1,
      currentCategory: 'All',
      expiredQuestions: [],
      overallTimerExpired: false
    });
  }

  private handleOverallTimerExpired = (): void => {
    this.setState(
      {
        overallTimerExpired: true
      },
      async () => {
        try {
          await this.handleSubmitQuiz();
        } catch (error) {
          console.error("Error submitting quiz:", error);
        }
      }
    );
  }


  private handleAnswerSelect = (questionId: number, choiceId: string | string[]): void => {
    const updatedQuestions = [...this.state.questions];
    const questionIndex = updatedQuestions.findIndex(q => q.id === questionId);

    if (questionIndex !== -1) {
      updatedQuestions[questionIndex].selectedChoice = choiceId;

      // Count answered questions
      const answeredQuestions = updatedQuestions.filter(q => q.selectedChoice !== undefined).length;

      this.setState({
        questions: updatedQuestions,
        answeredQuestions
      });
    }
  }

  private handlePageChange = (page: number): void => {
    this.setState({ currentPage: page }, () => {
      const questionsContainer = document.querySelector(`.${styles.questionsContainer}`);
      if (questionsContainer) {
        questionsContainer.scrollIntoView({
          behavior: 'smooth',
          block: 'start'
        });
      } else {
        window.scrollTo({
          top: 0,
          behavior: 'smooth'
        });
      }
    });
  };
  private prepareDetailedResults = (
    questions: IQuizQuestion[],
    score: number,
    totalPoints: number
  ): IDetailedQuizResults => {
    // Track all questions
    const totalQuestions = questions.length;
    const answeredQuestions = questions.filter(q => q.selectedChoice !== undefined).length;
    const correctlyAnsweredQuestions = questions.filter(q =>
      q.selectedChoice !== undefined && this.isQuestionCorrect(q)
    ).length;

    // Calculate percentages based on ALL questions
    const percentageAnswered = Math.round((answeredQuestions / totalQuestions) * 100);
    const percentageCorrect = Math.round((correctlyAnsweredQuestions / totalQuestions) * 100);
    const percentageCorrectOfAnswered = answeredQuestions > 0
      ? Math.round((correctlyAnsweredQuestions / answeredQuestions) * 100)
      : 0;

    // Prepare question-by-question results
    const questionResults: IQuestionResult[] = questions.map(question => {
      // Handle unanswered questions
      if (question.selectedChoice === undefined) {
        return {
          id: question.id,
          title: question.title,
          userAnswer: undefined,
          correctAnswer: this.getCorrectAnswerText(question),
          isCorrect: false,
          points: question.points || 1,
          earnedPoints: 0,
          explanation: question.explanation
        };
      }

      // Determine if the question was answered correctly
      const isCorrect = this.isQuestionCorrect(question);
      const points = question.points || 1;

      // Format user answer for display
      let formattedUserAnswer: string | string[] | undefined;

      switch (question.type) {
        case QuestionType.MultipleChoice:
        case QuestionType.TrueFalse: {
          const selectedChoice = question.choices.find(c => c.id === question.selectedChoice);
          formattedUserAnswer = selectedChoice ? selectedChoice.text : undefined;
          break;
        }

        case QuestionType.MultiSelect: {
          if (Array.isArray(question.selectedChoice)) {
            formattedUserAnswer = question.choices
              .filter(choice => question.selectedChoice?.includes(choice.id))
              .map(choice => choice.text);
          }
          break;
        }

        case QuestionType.Matching: {
          if (Array.isArray(question.selectedChoice)) {
            const userSelections = question.selectedChoice.map(selection => {
              if (typeof selection === 'string') {
                const [leftId, rightId] = selection.split(':');
                const leftItem = question.matchingPairs?.find(p => p.id === leftId)?.leftItem;
                const rightItem = question.matchingPairs?.find(p => p.id === rightId)?.rightItem;
                return leftItem && rightItem ? `${leftItem} → ${rightItem}` : '';
              }
              return '';
            }).filter(Boolean);

            formattedUserAnswer = userSelections.length > 0 ? userSelections : undefined;
          }
          break;
        }

        case QuestionType.ShortAnswer:
          formattedUserAnswer = question.selectedChoice as string;
          break;
      }

      return {
        id: question.id,
        title: question.title,
        userAnswer: formattedUserAnswer,
        correctAnswer: this.getCorrectAnswerText(question),
        isCorrect,
        points,
        earnedPoints: isCorrect ? points : 0,
        explanation: question.explanation
      };
    });

    return {
      score,
      totalPoints,
      totalQuestions,
      answeredQuestions,
      correctlyAnsweredQuestions,
      percentageAnswered,
      percentageCorrect,
      percentageCorrectOfAnswered,
      percentage: Math.round((score / totalPoints) * 100),
      questionResults,
      timestamp: new Date().toISOString()
    };
  };


  private handleSubmitQuiz = async (): Promise<void> => {
    this.setState({
      isSubmitting: true,
      submissionError: ''
    });

    try {
      // Calculate points for ALL questions
      let totalQuizPoints = 0;
      let earnedPoints = 0;
      const allQuestions = this.state.questions;

      // Calculate total possible points and earned points
      allQuestions.forEach(question => {
        // Get points for this question (default to 1 if not specified)
        const questionPoints = question.points || 1;
        totalQuizPoints += questionPoints;

        // Only add points if the question was answered and correct
        if (question.selectedChoice !== undefined && this.isQuestionCorrect(question)) {
          earnedPoints += questionPoints;
        }
      });

      // Save results to SharePoint list
      const savedSuccessfully = await this.saveQuizResults();

      // Prepare detailed results with comprehensive metrics
      const detailedResults = this.prepareDetailedResults(
        allQuestions,
        earnedPoints,
        totalQuizPoints
      );

      // Log the results for debugging
      console.log('Quiz submission results:', {
        totalQuizPoints,
        earnedPoints,
        savedSuccessfully
      });

      // Delete any existing saved progress
      if (this.state.savedProgressId) {
        await this.deleteSavedProgress(this.state.savedProgressId);
      }

      // Update state with results
      this.setState({
        showResults: true,
        score: earnedPoints,
        totalQuestions: allQuestions.length,
        totalPoints: totalQuizPoints,
        answeredQuestions: allQuestions.filter(q => q.selectedChoice !== undefined).length,
        submissionSuccess: savedSuccessfully,
        isSubmitting: false,
        detailedResults,

        // Clear any saved progress state
        savedProgressId: undefined,
        hasSavedProgress: false
      });

    } catch (error) {
      console.error('Error submitting quiz:', error);

      this.setState({
        submissionError: this.props.errorMessage || 'An error occurred while submitting your quiz.',
        isSubmitting: false
      });

      // Optional: Delete saved progress even if submission fails
      if (this.state.savedProgressId) {
        try {
          await this.deleteSavedProgress(this.state.savedProgressId);
        } catch (deleteError) {
          console.error('Error deleting saved progress during failed submission:', deleteError);
        }
      }
    }
  };



  private handleRetakeQuiz = (): void => {
    const resetQuestions = this.state.originalQuestions.map(q => ({
      ...q,
      selectedChoice: undefined,
      userAnswer: undefined,
      isCorrect: undefined
    }));

    const finalQuestions = this.props.randomizeQuestions
      ? this.shuffleArray(resetQuestions)
      : resetQuestions;

    const finalQuestionsWithRandomAnswers = this.props.randomizeAnswers
      ? finalQuestions.map(question => ({
        ...question,
        choices: this.shuffleArray([...question.choices])
      }))
      : finalQuestions;

    this.setState({
      questions: finalQuestionsWithRandomAnswers,
      currentPage: 1,
      currentCategory: 'All',
      showResults: false,
      score: 0,
      totalQuestions: 0,
      totalPoints: 0,
      answeredQuestions: 0,
      submissionSuccess: false,
      submissionError: '',
      showStartPage: true,
      quizStarted: false,
      overallTimerExpired: false,
      expiredQuestions: [],
      savedProgressId: undefined,
      hasSavedProgress: false,
      detailedResults: undefined
    });

    if (this.props.timeLimit) {
      this.setRemainingTime(this.props.timeLimit);
    }
  };

  private handleAddQuestion = (): void => {
    this.setState({
      showAddQuestionForm: true,
      newQuestion: this.getEmptyQuestion()
    });
  }
  private handleEditQuestion = (question: IQuizQuestion): void => {
    this.setState({
      showAddQuestionForm: true,
      newQuestion: question
    });
  }


  private handleAddQuestionSubmit = (newQuestion: IQuizQuestion): void => {
    let updatedQuestions: IQuizQuestion[];

    const existingQuestionIndex = this.props.questions.findIndex(q => q.id === newQuestion.id);

    if (existingQuestionIndex >= 0) {
      updatedQuestions = this.props.questions.map(q =>
        q.id === newQuestion.id ? newQuestion : q
      );
    } else {
      updatedQuestions = [...this.props.questions, newQuestion];
    }

    this.props.updateQuestions(updatedQuestions);

    this.setState({
      showAddQuestionForm: false,
      questions: updatedQuestions,
      originalQuestions: updatedQuestions,
      categories: this.updateCategories(updatedQuestions)
    });
  }


  private updateCategories(questions: IQuizQuestion[]): string[] {
    const categoriesSet = new Set<string>();
    questions.forEach(q => {
      if (q.category) categoriesSet.add(q.category);
    });
    return ['All', ...Array.from(categoriesSet)];
  }

  private handleAddQuestionCancel = (): void => {
    this.setState({ showAddQuestionForm: false });
  }

  private handleImportQuestions = (): void => {
    this.setState({ importDialogOpen: true });
  }

  private handleImportQuestionsSubmit = (importedQuestions: IQuizQuestion[]): void => {
    const updatedQuestions = [...this.props.questions, ...importedQuestions];

    // Update questions through the prop callback
    this.props.updateQuestions(updatedQuestions);

    // Update local state for immediate re-render
    this.setState({
      importDialogOpen: false,
      questions: updatedQuestions,
      originalQuestions: updatedQuestions,
      categories: this.updateCategories(updatedQuestions)
    });
  }



  private handleImportQuestionsCancel = (): void => {
    this.setState({ importDialogOpen: false });
  }

  private handleExportQuestions = (): void => {
    const dataStr = JSON.stringify(this.props.questions, null, 2);
    const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(dataStr);
    const exportFileDefaultName = `quiz-questions-${new Date().toISOString().slice(0, 10)}.json`;

    // Create a download link and trigger the download
    const linkElement = document.createElement('a');
    linkElement.setAttribute('href', dataUri);
    linkElement.setAttribute('download', exportFileDefaultName);
    linkElement.style.display = 'none';
    document.body.appendChild(linkElement);
    linkElement.click();
    document.body.removeChild(linkElement);
  }


  private handlePreviewQuestion = (question: IQuizQuestion): void => {
    this.setState({
      previewQuestion: question,
      showQuestionPreview: true
    });
  }

  private handleDeleteQuestion = (questionId: number): void => {
    // Filter out just the question with the matching ID
    const updatedQuestions = this.props.questions.filter(q => q.id !== questionId);

    // Update questions through the prop callback
    this.props.updateQuestions(updatedQuestions);

    // Update local state
    this.setState({
      questions: updatedQuestions,
      originalQuestions: updatedQuestions,
      categories: this.updateCategories(updatedQuestions)
    });
  }




  private executeConfirmedAction = (): void => {
    const { confirmDialogAction, previewQuestion } = this.state;
    let updatedQuestions = [...this.props.questions];

    if (confirmDialogAction === 'deleteQuestion' && previewQuestion) {
      updatedQuestions = this.props.questions.filter(q => q.id !== previewQuestion.id);
    } else if (confirmDialogAction === 'deleteAllQuestions') {
      updatedQuestions = [];
    }

    // Update questions through the prop callback
    this.props.updateQuestions(updatedQuestions);

    // Update local state for immediate re-render
    this.setState({
      showConfirmDialog: false,
      confirmDialogAction: '',
      previewQuestion: undefined,
      questions: updatedQuestions,
      originalQuestions: updatedQuestions,
      categories: this.updateCategories(updatedQuestions)
    });
  }



  private handleClosePreview = (): void => {
    this.setState({ showQuestionPreview: false });
  }


  private handleSubmitRequireAllChange = (
    ev?: React.FormEvent<HTMLElement | HTMLInputElement>,
    checked?: boolean
  ): void => {
    this.setState({ submitRequireAllAnswered: !!checked });
  }

  // Render methods
  private renderAdminPanel(): JSX.Element | null {
    if (this.props.displayMode !== DisplayMode.Edit) return null;

    return (
      <Stack horizontal tokens={stackTokens} className={styles.adminPanel}>
        <div className={styles.buttonGroup}>
          <PrimaryButton
            iconProps={addIcon}
            text="Add Question"
            onClick={this.handleAddQuestion}
            className={`${styles.actionButton} ${styles.primary}`}
          />
          <DefaultButton
            iconProps={editIcon}
            text="Edit Questions"
            onClick={this.handleOpenEditQuestionsDialog}
            className={`${styles.actionButton} ${styles.secondary}`}
          />
          <DefaultButton
            iconProps={importIcon}
            text="Import"
            onClick={this.handleImportQuestions}
            className={`${styles.actionButton} ${styles.secondary}`}
          />
          <DefaultButton
            iconProps={exportIcon}
            text="Export"
            onClick={this.handleExportQuestions}
            className={`${styles.actionButton} ${styles.secondary}`}
          />

        </div>
      </Stack>
    );
  }




  private renderConfirmDialog(): JSX.Element {
    const { showConfirmDialog, confirmDialogAction, previewQuestion } = this.state;

    let dialogTitle = 'Confirm Action';
    let dialogMessage = 'Are you sure you want to perform this action?';

    if (confirmDialogAction === 'deleteQuestion' && previewQuestion) {
      dialogTitle = 'Confirm Delete Question';
      dialogMessage = `Are you sure you want to delete the question "${previewQuestion.title}"?`;
    } else if (confirmDialogAction === 'deleteAllQuestions') {
      dialogTitle = 'Confirm Delete All Questions';
      dialogMessage = 'Are you sure you want to delete all questions? This action cannot be undone.';
    }

    return (
      <Dialog
        hidden={!showConfirmDialog}
        onDismiss={() => this.setState({ showConfirmDialog: false })}
        dialogContentProps={{
          type: DialogType.normal,
          title: dialogTitle,
          subText: dialogMessage
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 450 } }
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={this.executeConfirmedAction} text="Yes" />
          <DefaultButton onClick={() => this.setState({ showConfirmDialog: false })} text="No" />
        </DialogFooter>
      </Dialog>
    );
  }


  public render(): React.ReactElement<IQuizProps> {
    const {
      loading,
      questions,
      categories,
      currentCategory,
      currentPage,
      showResults,
      submissionError,
      isSubmitting,
      answeredQuestions,
      submitRequireAllAnswered,
      showAddQuestionForm,
      importDialogOpen,
      showQuestionPreview,
      previewQuestion,
      showEditQuestionsDialog,
      detailedResults,
      showStartPage,
      quizStarted,
      overallTimerExpired
    } = this.state;

    const {
      questionsPerPage,
      showProgressIndicator,
      displayMode,
    } = this.props;

    if (questions.length === 0) {
      return (
        <Stack styles={mainContainerStyles}>
          <WebPartTitle
            displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
          />

          {this.renderAdminPanel()}

          <div className={styles.emptyState}>
            <Text variant="large">No questions have been added to this quiz yet.</Text>
            <Text>Use the admin panel to add questions or import from a file.</Text>

            {displayMode === DisplayMode.Edit && (
              <Stack horizontal tokens={stackTokens} horizontalAlign="center" style={{ marginTop: '16px' }}>
                <PrimaryButton
                  text="Add First Question"
                  onClick={this.handleAddQuestion}
                  iconProps={addIcon}
                />
              </Stack>
            )}
          </div>

          {/* Dialogs */}
          {showAddQuestionForm && (
            <AddQuestionDialog
              categories={categories.filter(cat => cat !== 'All')}
              onSubmit={this.handleAddQuestionSubmit}
              onCancel={this.handleAddQuestionCancel}
              isSubmitting={false}
              onPreviewQuestion={this.handlePreviewQuestion}
              context={this.props.context}
            />
          )}

          {importDialogOpen && (
            <ImportQuestionsDialog
              existingCategories={categories.filter(cat => cat !== 'All')}
              onImportQuestions={this.handleImportQuestionsSubmit}
              onCancel={this.handleImportQuestionsCancel}
            />
          )}

          {this.renderConfirmDialog()}
          {this.renderSaveProgressDialog()}
          {this.renderResumeDialog()}
        </Stack>
      );
    }

    if (displayMode === DisplayMode.Read && showStartPage) {
      return (
        <Stack styles={mainContainerStyles}>
          <WebPartTitle
            displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
          />

          <QuizStartPage
            title={this.props.title}
            onStartQuiz={this.handleStartQuiz}
            totalQuestions={questions.length}
            totalPoints={questions.reduce((sum, q) => sum + (q.points || 1), 0)}
            categories={categories.filter(c => c !== 'All')}
            timeLimit={this.props.timeLimit}
            passingScore={this.props.passingScore}
            description="This quiz will test your knowledge of the subject matter. Please read each question carefully before selecting your answer."
            hasSavedProgress={this.state.hasSavedProgress}
            onResumeQuiz={this.resumeQuiz}
          />
        </Stack>
      );
    }

    // Loading state
    if (loading) {
      return (
        <Stack styles={mainContainerStyles} horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
          <Spinner size={SpinnerSize.large} label="Loading quiz..." />
        </Stack>
      );
    }

    // Results view
    if (showResults) {
      return (
        <Stack styles={mainContainerStyles}>
          <WebPartTitle
            displayMode={this.props.displayMode}
            title={this.props.title}
            updateProperty={this.props.updateProperty}
          />

          <QuizResults
            score={this.state.score}
            totalQuestions={this.state.totalQuestions}
            totalPoints={this.state.totalPoints}
            isSubmitting={isSubmitting}
            submissionSuccess={this.state.submissionSuccess}
            submissionError={submissionError}
            onRetakeQuiz={this.handleRetakeQuiz}
            messages={{
              excellent: this.props.excellentScoreMessage,
              good: this.props.goodScoreMessage,
              average: this.props.averageScoreMessage,
              poor: this.props.poorScoreMessage,
              success: this.props.resultsSavedMessage
            }}
            detailedResults={detailedResults}
          />
        </Stack>
      );
    }

    // Filter questions by category
    const filteredQuestions = currentCategory === 'All'
      ? questions
      : questions.filter(q => q.category === currentCategory);

    // Paginate questions
    const startIndex = (currentPage - 1) * questionsPerPage;
    const paginatedQuestions = filteredQuestions.slice(startIndex, startIndex + questionsPerPage);

    const allQuestionsAnswered = questions.length === answeredQuestions;
    const submitEnabled = !submitRequireAllAnswered ? answeredQuestions > 0 : allQuestionsAnswered;

    // Quiz taker view (or edit mode)
    return (
      <Stack styles={mainContainerStyles}>
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty}
        />

        {this.renderAdminPanel()}

        {submissionError && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            styles={{ root: { marginBottom: 15 } }}
          >
            {submissionError}
          </MessageBar>
        )}

        {/* Overall Quiz Timer - Only show when quiz is started in display mode */}
        {displayMode === DisplayMode.Read && quizStarted && this.props.timeLimit && this.props.timeLimit > 0 && (
          <QuizTimer
            timeLimit={this.props.timeLimit}
            onTimeExpired={this.handleOverallTimerExpired}
            paused={showResults}
          />
        )}

        {overallTimerExpired && (
          <MessageBar
            messageBarType={MessageBarType.severeWarning}
            isMultiline={true}
            styles={{ root: { marginBottom: 15 } }}
          >
            The time limit for this quiz has expired. Your answers have been automatically submitted.
          </MessageBar>
        )}

        {displayMode === DisplayMode.Edit && (
          <Stack horizontal horizontalAlign="space-between" tokens={stackTokens}>
            <Checkbox
              label="Require all questions to be answered"
              checked={submitRequireAllAnswered}
              onChange={this.handleSubmitRequireAllChange}
            />
          </Stack>
        )}

        {/* Progress Indicator - New feature */}
        {showProgressIndicator && (
          <QuizProgressTracker
            progress={{
              currentQuestion: currentPage,
              totalQuestions: filteredQuestions.length,
              answeredQuestions,
              percentage: filteredQuestions.length > 0
                ? Math.round((answeredQuestions / filteredQuestions.length) * 100)
                : 0
            }}
            showPercentage={true}
            showNumbers={true}
            showIcon={true}
          />
        )}

        <Pivot
          selectedKey={currentCategory}
          onLinkClick={this.handleCategoryChange}
          className={styles.categoryFilter}
        >
          {categories.map(category => (
            <PivotItem key={category} headerText={category} itemKey={category} />
          ))}
        </Pivot>

        {filteredQuestions.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No questions found for this category.
          </MessageBar>
        ) : (
          <>
            <div className={styles.questionsContainer}>
              {paginatedQuestions.map((question, index) => (
                <QuizQuestion
                  key={question.id}
                  question={question}
                  onAnswerSelect={this.handleAnswerSelect}
                  questionNumber={startIndex + index + 1}
                  totalQuestions={filteredQuestions.length}
                  showProgressIndicator={false}
                  onTimeExpired={(questionId) => this.handleQuestionTimeExpired(questionId)}
                />
              ))}
            </div>

            {filteredQuestions.length > questionsPerPage && (
              <div className={styles.paginationContainer}>
                <Pagination
                  currentPage={currentPage}
                  totalPages={Math.ceil(filteredQuestions.length / questionsPerPage)}
                  onChange={this.handlePageChange}
                  limiter={3}
                />
              </div>
            )}

            <div className={styles.submitContainer}>
              {isSubmitting ? (
                <Spinner size={SpinnerSize.small} label="Submitting quiz..." />
              ) : (
                <>
                  <PrimaryButton
                    iconProps={submitIcon}
                    text="Submit Quiz"
                    onClick={this.handleSubmitQuiz}
                    disabled={!submitEnabled}
                  />
                  <DefaultButton
                    iconProps={saveIcon}
                    text="Save & Continue Later"
                    onClick={this.handleSaveProgressClick}
                    style={{ marginLeft: '10px' }}
                  />
                </>
              )}
              {!submitEnabled && submitRequireAllAnswered && (
                <Text style={{ marginTop: '8px', color: '#a4262c' }}>
                  Please answer all questions before submitting.
                </Text>
              )}
            </div>
          </>
        )}

        {/* Dialogs */}
        {showQuestionPreview && previewQuestion && (
          <QuestionPreview
            question={previewQuestion}
            onClose={this.handleClosePreview}
          />
        )}

        {showEditQuestionsDialog && (
          <Dialog
            hidden={false}
            onDismiss={this.handleCloseEditQuestionsDialog}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: "Manage Quiz Questions",
              className: styles.dialogContent
            }}
            modalProps={{
              isBlocking: false,
              className: styles.editQuestionsDialog
            }}
          >
            <div className={styles.questionsContainer}>
              {questions.length === 0 ? (
                <div className={styles.emptyState}>
                  <Text>No questions have been added yet.</Text>
                  <Stack horizontal horizontalAlign="center" tokens={stackTokens} style={{ marginTop: '16px' }}>
                    <PrimaryButton
                      text="Add First Question"
                      onClick={() => {
                        this.setState({
                          showAddQuestionForm: true,
                          showEditQuestionsDialog: false
                        });
                      }}
                      iconProps={addIcon}
                    />
                  </Stack>
                </div>
              ) : (
                <>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: '16px' }}>
                    <Text variant="large">{questions.length} questions total</Text>
                  </Stack>

                  <QuestionManagement
                    questions={questions}
                    onUpdateQuestions={(updatedQuestions) => {
                      // Update questions through the prop callback
                      this.props.updateQuestions(updatedQuestions);

                      // Update local state for immediate re-render
                      this.setState({
                        questions: updatedQuestions,
                        originalQuestions: updatedQuestions,
                        categories: this.updateCategories(updatedQuestions)
                      });
                    }}
                    onAddQuestion={() => {
                      this.setState({
                        showAddQuestionForm: true,
                        showEditQuestionsDialog: false
                      });
                    }}
                    onEditQuestion={(question) => {
                      this.setState({
                        showAddQuestionForm: true,
                        newQuestion: question,
                        showEditQuestionsDialog: false
                      });
                    }}
                    onPreviewQuestion={(question) => {
                      this.setState({
                        previewQuestion: question,
                        showQuestionPreview: true
                      });
                    }}
                    onDeleteQuestion={this.handleDeleteQuestion}
                  />
                </>
              )}
            </div>
            <DialogFooter>
              <PrimaryButton onClick={this.handleCloseEditQuestionsDialog} text="Close" />
            </DialogFooter>
          </Dialog>
        )}

        {showAddQuestionForm && (
          <AddQuestionDialog
            categories={categories.filter(cat => cat !== 'All')}
            onSubmit={this.handleAddQuestionSubmit}
            onCancel={this.handleAddQuestionCancel}
            isSubmitting={false}
            onPreviewQuestion={this.handlePreviewQuestion}
            initialQuestion={this.state.newQuestion.id !== Date.now() ? this.state.newQuestion : undefined}
            context={this.props.context}
          />
        )}

        {importDialogOpen && (
          <ImportQuestionsDialog
            existingCategories={categories.filter(cat => cat !== 'All')}
            onImportQuestions={this.handleImportQuestionsSubmit}
            onCancel={this.handleImportQuestionsCancel}
          />
        )}

        {this.renderConfirmDialog()}
        {this.renderSaveProgressDialog()}
        {this.renderResumeDialog()}
      </Stack>
    );
  }
}
