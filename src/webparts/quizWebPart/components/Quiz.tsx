import * as React from 'react';
import { IQuizProps } from './IQuizProps';
import { IQuizState, IQuizQuestion, QuestionType } from './interfaces';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { v4 as uuidv4 } from 'uuid';
import QuizQuestion from './QuizQuestion';
import QuizResults from './QuizResults';
import AddQuestionDialog from './AddQuestionDialog';
import ImportQuestionsDialog from './ImportQuestionsDialog';
import QuestionPreview from './QuestionPreview';
import { DisplayMode } from '@microsoft/sp-core-library';

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
  ProgressIndicator,
  Dialog,
  DialogType,
  DialogFooter,
  mergeStyles,
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

// Styles
const mainContainerStyles: IStackStyles = {
  root: {
    padding: '20px',
    maxWidth: '1200px',
    margin: '0 auto'
  }
};

const progressIndicatorClass = mergeStyles({
  marginBottom: '16px'
});

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
      showEditQuestionsDialog: false

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
      ]
    };
  }

  public componentDidMount(): void {
    this.randomizeQuestionsIfNeeded();
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
  private saveQuizResults = async (): Promise<boolean> => {
    try {
      const spHttpClient = this.props.context.spHttpClient;
      const webUrl = this.props.context.pageContext.web.absoluteUrl;
      const currentUser = this.props.context.pageContext.user;

      // Prepare detailed result data
      const resultData = {
        Title: `Quiz Result - ${new Date().toISOString()}`,
        UserName: currentUser.displayName,
        UserEmail: currentUser.email,
        UserId: currentUser.loginName,
        QuizTitle: this.props.title,

        // Score details
        Score: this.state.score,
        TotalPoints: this.state.totalPoints,
        ScorePercentage: Math.round((this.state.score / this.state.totalPoints) * 100),

        // Timestamp
        ResultDate: new Date().toISOString(),

        // Detailed question results
        QuestionDetails: JSON.stringify(
          this.state.questions.map(question => ({
            QuestionId: question.id,
            QuestionTitle: question.title,
            QuestionType: question.type,
            SelectedChoice: question.selectedChoice,
            IsCorrect: this.isQuestionCorrect(question),
            Points: question.points || 1
          }))
        )
      };

      // Check if list exists, create if not
      const listExists = await this.ensureQuizResultsList(spHttpClient, webUrl);
      if (!listExists) {
        throw new Error('Could not create or find Quiz Results list');
      }

      // Save result
      const response = await spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('QuizResults')/items`,
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
        throw new Error(`Failed to save quiz results: ${errorText}`);
      }

      return true;
    } catch (error) {
      console.error('Error saving quiz results:', error);

      // Update state with submission error
      this.setState({
        submissionError: error instanceof Error
          ? error.message
          : 'An unexpected error occurred while saving results'
      });

      return false;
    }
  }

  // Helper method to ensure Quiz Results list exists
  private ensureQuizResultsList = async (
    spHttpClient: SPHttpClient,
    webUrl: string
  ): Promise<boolean> => {
    try {
      // First, check if list exists
      const listCheckResponse = await spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('QuizResults')`,
        SPHttpClient.configurations.v1
      );

      if (listCheckResponse.ok) {
        return true;
      }

      // If list doesn't exist, create it
      const listCreationResponse = await spHttpClient.post(
        `${webUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'BaseTemplate': 100, // Generic list
            'Title': 'QuizResults',
            'Description': 'Stores quiz results for tracking and analysis'
          })
        }
      );

      if (!listCreationResponse.ok) {
        const errorText = await listCreationResponse.text();
        throw new Error(`Failed to create QuizResults list: ${errorText}`);
      }

      // Add columns to the list
      await this.addListColumns(spHttpClient, webUrl);

      return true;
    } catch (error) {
      console.error('Error ensuring QuizResults list:', error);
      return false;
    }
  }

  // Add necessary columns to the list
  private addListColumns = async (
    spHttpClient: SPHttpClient,
    webUrl: string
  ): Promise<void> => {
    const columnDefinitions = [
      {
        Title: 'UserName',
        Type: 'Text'
      },
      {
        Title: 'UserEmail',
        Type: 'Text'
      },
      {
        Title: 'UserId',
        Type: 'Text'
      },
      {
        Title: 'QuizTitle',
        Type: 'Text'
      },
      {
        Title: 'Score',
        Type: 'Number'
      },
      {
        Title: 'TotalPoints',
        Type: 'Number'
      },
      {
        Title: 'ScorePercentage',
        Type: 'Number'
      },
      {
        Title: 'ResultDate',
        Type: 'DateTime'
      },
      {
        Title: 'QuestionDetails',
        Type: 'Note'
      }
    ];

    for (const column of columnDefinitions) {
      try {
        await spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('QuizResults')/fields`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata'
            },
            body: JSON.stringify({
              '__metadata': { 'type': 'SP.Field' },
              'Title': column.Title,
              'FieldTypeKind': this.getFieldTypeKind(column.Type)
            })
          }
        );
      } catch (error) {
        console.warn(`Could not add column ${column.Title}:`, error);
      }
    }
  }

  // Helper to map field types
  private getFieldTypeKind = (type: string): number => {
    switch (type) {
      case 'Text': return 2;     // SP.FieldType.Text
      case 'Number': return 9;   // SP.FieldType.Number
      case 'DateTime': return 4; // SP.FieldType.DateTime
      case 'Note': return 3;     // SP.FieldType.Note
      default: return 2;         // Default to Text
    }
  }

  // Helper to determine if a question is correct
  private isQuestionCorrect = (question: IQuizQuestion): boolean => {
    if (question.selectedChoice === undefined) return false;
  
    switch (question.type) {
      case QuestionType.MultipleChoice:
      case QuestionType.TrueFalse: {
        // Wrap in block to create a new lexical scope
        const isCorrect = question.choices.find(
          c => c.id === question.selectedChoice && c.isCorrect
        ) !== undefined;
        return isCorrect;
      }
  
      case QuestionType.MultiSelect: {
        // Wrap in block to create a new lexical scope
        if (!Array.isArray(question.selectedChoice)) return false;
        const selectedIds = new Set(question.selectedChoice);
        const correctIds = new Set(
          question.choices.filter(c => c.isCorrect).map(c => c.id)
        );
        return selectedIds.size === correctIds.size && 
               [...correctIds].every(id => selectedIds.has(id));
      }
  
      case QuestionType.ShortAnswer: {
        // Wrap in block to create a new lexical scope
        const isCorrect = question.correctAnswer !== undefined && 
          (question.caseSensitive 
            ? question.selectedChoice === question.correctAnswer
            : (question.selectedChoice as string).toLowerCase() === 
              (question.correctAnswer as string).toLowerCase());
        return isCorrect;
      }
  
      default:
        return false;
    }
  }
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
    this.setState({ currentPage: page });
  }

  private handleSubmitQuiz = async (): Promise<void> => {
    this.setState({ isSubmitting: true });

    try {
      // Calculate score
      let score = 0;
      let totalPoints = 0;
      let answeredQuestions = 0;
      const allQuestions = this.state.questions;
      const savedSuccessfully = await this.saveQuizResults();

      allQuestions.forEach(question => {
        if (question.selectedChoice !== undefined) {
          answeredQuestions++;
          const points = question.points || 1;
          totalPoints += points;

          let isCorrect = false;

          switch (question.type) {
            case QuestionType.MultipleChoice:
            case QuestionType.TrueFalse: {
              const selectedChoice = question.choices.find(c => c.id === question.selectedChoice);
              isCorrect = !!selectedChoice?.isCorrect;
              break;
            }
            case QuestionType.MultiSelect: {
              if (Array.isArray(question.selectedChoice)) {
                const selectedIds = new Set(question.selectedChoice);
                const correctChoiceIds = question.choices.filter(c => c.isCorrect).map(c => c.id);
                isCorrect = correctChoiceIds.length > 0 &&
                  correctChoiceIds.every(id => selectedIds.has(id)) &&
                  selectedIds.size === correctChoiceIds.length;
              }
              break;
            }
            case QuestionType.ShortAnswer: {
              if (typeof question.selectedChoice === 'string' && question.correctAnswer) {
                isCorrect = question.caseSensitive
                  ? question.selectedChoice.trim() === question.correctAnswer.trim()
                  : question.selectedChoice.trim().toLowerCase() === question.correctAnswer.trim().toLowerCase();
              }
              break;
            }
          }

          if (isCorrect) {
            score += points;
          }
        }
      });

      await new Promise(resolve => setTimeout(resolve, 1000));

      this.setState({
        showResults: true,
        score,
        totalQuestions: answeredQuestions,
        totalPoints,
        submissionSuccess: savedSuccessfully,
        isSubmitting: false
      });
    } catch (error) {
      console.error('Error submitting quiz:', error);
      this.setState({
        submissionError: this.props.errorMessage || 'An error occurred while submitting your quiz.',
        isSubmitting: false
      });
    }
  }

  private handleRetakeQuiz = (): void => {
    const resetQuestions = this.state.questions.map(q => ({
      ...q,
      selectedChoice: undefined
    }));

    this.setState({
      questions: resetQuestions,
      currentPage: 1,
      currentCategory: 'All',
      showResults: false,
      score: 0,
      totalPoints: 0,
      answeredQuestions: 0,
      submissionSuccess: false,
      submissionError: ''
    });

    this.randomizeQuestionsIfNeeded();
  }

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
    // Create a new array with the new question
    const updatedQuestions = [...this.props.questions, newQuestion];

    // Update questions through the prop callback
    this.props.updateQuestions(updatedQuestions);

    // Update local state
    this.setState({
      showAddQuestionForm: false,
      questions: updatedQuestions,  // Immediately update local state
      originalQuestions: updatedQuestions,  // Update original questions too
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
    const updatedQuestions = this.props.questions.filter(q => q.id !== questionId);
    this.props.updateQuestions(updatedQuestions);
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


  // Updated handler signature to fix Checkbox onChange type error
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

  private renderProgressBar(): JSX.Element {
    const { questions, answeredQuestions } = this.state;
    const totalQuestions = questions.length;
    const percentage = totalQuestions > 0 ? (answeredQuestions / totalQuestions) * 100 : 0;

    return (
      <div className={progressIndicatorClass}>
        <Text>{`${answeredQuestions} of ${totalQuestions} questions answered (${Math.round(percentage)}%)`}</Text>
        <ProgressIndicator percentComplete={percentage / 100} />
      </div>
    );
  }

  // Update the render return type to allow null
  public render(): React.ReactElement<IQuizProps> | undefined {
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
      showEditQuestionsDialog
    } = this.state;

    const { questionsPerPage, showProgressIndicator, displayMode } = this.props;

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

          {showAddQuestionForm && (
            <AddQuestionDialog
              categories={categories.filter(cat => cat !== 'All')}
              onSubmit={this.handleAddQuestionSubmit}
              onCancel={this.handleAddQuestionCancel}
              isSubmitting={false}
              onPreviewQuestion={this.handlePreviewQuestion}
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

    if (loading) {
      return (
        <Stack styles={mainContainerStyles} horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
          <Spinner size={SpinnerSize.large} label="Loading quiz..." />
        </Stack>
      );
    }

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
          />
        </Stack>
      );
    }

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

        {displayMode === DisplayMode.Edit && (
          <Stack horizontal horizontalAlign="space-between" tokens={stackTokens}>
            <Checkbox
              label="Require all questions to be answered"
              checked={submitRequireAllAnswered}
              onChange={this.handleSubmitRequireAllChange}
            />
          </Stack>
        )}

        {showProgressIndicator && this.renderProgressBar()}

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
                <PrimaryButton
                  iconProps={submitIcon}
                  text="Submit Quiz"
                  onClick={this.handleSubmitQuiz}
                  disabled={!submitEnabled}
                />
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
      </Stack>
    );
  }
}
