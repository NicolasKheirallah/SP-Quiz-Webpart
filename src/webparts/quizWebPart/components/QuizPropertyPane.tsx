import * as React from 'react';
import { useState, useMemo } from 'react';
import { 
  IQuizQuestion, 
  QuestionType 
} from './interfaces';
import QuestionManagement from './QuestionManagement';
import AddQuestionDialog from './AddQuestionDialog';
import ImportQuestionsDialog from './ImportQuestionsDialog';
import QuestionPreview from './QuestionPreview';
import {
  // Fluent UI React (legacy)
  Dropdown as FluentDropdown,
  IDropdownOption,
  Dialog as FluentDialog,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  MessageBar as FluentMessageBar,
  MessageBarType,
  IconButton,
  IIconProps,
  Stack,
  Text as FluentText,
  mergeStyles
} from '@fluentui/react';
import { DialogType as FluentDialogType } from '@fluentui/react';
interface IQuizPropertyPaneProps {
  questions: IQuizQuestion[];
  onUpdateQuestions: (questions: IQuizQuestion[]) => void;
}

// Rename this enum to avoid conflict with imported DialogType
enum DialogTypeEnum {
  None,
  AddQuestion,
  EditQuestion,
  ImportQuestions,
  Preview,
  BulkDelete,
  Filter,
  Sort
}

// Icons
const addIcon: IIconProps = { iconName: 'Add' };
const importIcon: IIconProps = { iconName: 'Download' };
const shuffleIcon: IIconProps = { iconName: 'Refresh' };
const filterIcon: IIconProps = { iconName: 'Filter' };
const sortIcon: IIconProps = { iconName: 'Sort' };
const deleteIcon: IIconProps = { iconName: 'Delete' };

// Styles
const containerClass = mergeStyles({
  padding: '12px',
  fontFamily: '"Segoe UI", sans-serif'
});

const headerClass = mergeStyles({
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  marginBottom: '16px'
});

const titleClass = mergeStyles({
  fontSize: '18px',
  fontWeight: 600,
  margin: 0,
  padding: 0
});

const buttonGroupClass = mergeStyles({
  display: 'flex',
  gap: '8px',
  flexWrap: 'wrap'
});

const messageBarClass = mergeStyles({
  marginBottom: '12px'
});

const noQuestionsClass = mergeStyles({
  padding: '20px',
  textAlign: 'center',
  backgroundColor: '#f3f2f1',
  borderRadius: '2px',
  marginTop: '12px'
});

const QuizPropertyPane: React.FC<IQuizPropertyPaneProps> = (props) => { 
  const { questions, onUpdateQuestions } = props;
  const [dialogType, setDialogType] = useState<DialogTypeEnum>(DialogTypeEnum.None);
  const [selectedQuestion, setSelectedQuestion] = useState<IQuizQuestion | undefined>(undefined);
  const [sortOption, setSortOption] = useState<string>('');
  const [filterOption, setFilterOption] = useState<string>('');
  const [successMessage, setSuccessMessage] = useState<string>('');

  // Sorting options
  const sortOptions: IDropdownOption[] = [
    { key: 'titleAsc', text: 'Title (A-Z)' },
    { key: 'titleDesc', text: 'Title (Z-A)' },
    { key: 'categoryAsc', text: 'Category (A-Z)' },
    { key: 'categoryDesc', text: 'Category (Z-A)' }
  ];

  // Filtering options
  const filterOptions: IDropdownOption[] = [
    { key: 'all', text: 'All Questions' },
    ...Object.values(QuestionType).map(type => ({ 
      key: type, 
      text: type.charAt(0).toUpperCase() + type.slice(1) 
    }))
  ];

  // Memoized categories
  const categories = useMemo(() => 
    Array.from(new Set(questions.map(q => q.category))).filter(Boolean), 
    [questions]
  );

  // Apply sorting
  const sortedQuestions = useMemo(() => {
    const sorted = [...questions];
    switch (sortOption) {
      case 'titleAsc':
        return sorted.sort((a, b) => a.title.localeCompare(b.title));
      case 'titleDesc':
        return sorted.sort((a, b) => b.title.localeCompare(a.title));
      case 'categoryAsc':
        return sorted.sort((a, b) => a.category.localeCompare(b.category));
      case 'categoryDesc':
        return sorted.sort((a, b) => b.category.localeCompare(a.category));
      default:
        return sorted;
    }
  }, [questions, sortOption]);

  // Apply filtering
  const filteredQuestions = useMemo(() => {
    if (filterOption === 'all' || !filterOption) return sortedQuestions;
    return sortedQuestions.filter(q => q.type === filterOption);
  }, [sortedQuestions, filterOption]);
  
  // Event handlers for question management
  const handleAddQuestion = (): void => {
    setDialogType(DialogTypeEnum.AddQuestion);
  };
  
  const handleEditQuestion = (question: IQuizQuestion): void => {
    setSelectedQuestion(question);
    setDialogType(DialogTypeEnum.EditQuestion);
  };
  
  const handlePreviewQuestion = (question: IQuizQuestion): void => {
    setSelectedQuestion(question);
    setDialogType(DialogTypeEnum.Preview);
  };
  
  const handleImportQuestions = (): void => {
    setDialogType(DialogTypeEnum.ImportQuestions);
  };
  
  // For adding a new question
  const handleAddQuestionSubmit = (newQuestion: IQuizQuestion): void => {
    onUpdateQuestions([...questions, newQuestion]);
    setDialogType(DialogTypeEnum.None);
    setSuccessMessage('Question added successfully');
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccessMessage(''), 3000);
  };
  
  // For editing an existing question
  const handleEditQuestionSubmit = (updatedQuestion: IQuizQuestion): void => {
    const updatedQuestions = questions.map((q: IQuizQuestion) => 
      q.id === updatedQuestion.id ? updatedQuestion : q
    );
    onUpdateQuestions(updatedQuestions);
    setDialogType(DialogTypeEnum.None);
    setSelectedQuestion(undefined);
    setSuccessMessage('Question updated successfully');
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccessMessage(''), 3000);
  };
  
  // For importing questions
  const handleImportQuestionsSubmit = (importedQuestions: IQuizQuestion[]): void => {
    onUpdateQuestions([...questions, ...importedQuestions]);
    setDialogType(DialogTypeEnum.None);
    setSuccessMessage(`${importedQuestions.length} questions imported successfully`);
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccessMessage(''), 3000);
  };
  
  // For bulk deletion
  const handleBulkDelete = (): void => {
    onUpdateQuestions([]);
    setDialogType(DialogTypeEnum.None);
    setSuccessMessage('All questions deleted successfully');
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccessMessage(''), 3000);
  };
  
  // For single question deletion
  const handleSingleDeleteQuestion = (questionId: number): void => {
    const updatedQuestions = questions.filter((q: IQuizQuestion) => q.id !== questionId);
    onUpdateQuestions(updatedQuestions);
    setSuccessMessage('Question deleted successfully');
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccessMessage(''), 3000);
  };
  
  // For randomizing questions
  const handleRandomize = (): void => {
    const shuffled = [...questions].sort(() => 0.5 - Math.random());
    onUpdateQuestions(shuffled);
    setSuccessMessage('Questions randomized successfully');
    
    // Clear success message after 3 seconds
    setTimeout(() => setSuccessMessage(''), 3000);
  };
  
  // Render the appropriate dialog based on dialogType
  const renderActiveDialog = (): React.ReactNode => {
    switch (dialogType) {
      case DialogTypeEnum.AddQuestion:
        return (
          <AddQuestionDialog
            categories={categories}
            onSubmit={handleAddQuestionSubmit}
            onCancel={() => setDialogType(DialogTypeEnum.None)}
            isSubmitting={false}
            onPreviewQuestion={handlePreviewQuestion}
          />
        );
      
      case DialogTypeEnum.EditQuestion:
        return selectedQuestion && (
          <AddQuestionDialog
            categories={categories}
            initialQuestion={selectedQuestion}
            onSubmit={handleEditQuestionSubmit}
            onCancel={() => {
              setDialogType(DialogTypeEnum.None);
              setSelectedQuestion(undefined);
            }}
            isSubmitting={false}
            onPreviewQuestion={handlePreviewQuestion}
          />
        );
      
      case DialogTypeEnum.ImportQuestions:
        return (
          <ImportQuestionsDialog
            existingCategories={categories}
            onImportQuestions={handleImportQuestionsSubmit}
            onCancel={() => setDialogType(DialogTypeEnum.None)}
          />
        );
      
      case DialogTypeEnum.Preview:
        return selectedQuestion && (
          <QuestionPreview
            question={selectedQuestion}
            onClose={() => {
              setDialogType(DialogTypeEnum.None);
              setSelectedQuestion(undefined);
            }}
          />
        );
      
      case DialogTypeEnum.BulkDelete:
        return (
          <FluentDialog
            hidden={false}
            onDismiss={() => setDialogType(DialogTypeEnum.None)}
            dialogContentProps={{
              type: FluentDialogType.close,
              title: 'Confirm Bulk Delete',
              subText: `Are you sure you want to delete all ${questions.length} questions? This action cannot be undone.`
            }}
            modalProps={{
              isBlocking: true,
              styles: { main: { maxWidth: 450 } }
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={handleBulkDelete} text="Delete" />
              <DefaultButton onClick={() => setDialogType(DialogTypeEnum.None)} text="Cancel" />
            </DialogFooter>
          </FluentDialog>
        );
      
      case DialogTypeEnum.Filter:
        return (
          <FluentDialog
            hidden={false}
            onDismiss={() => setDialogType(DialogTypeEnum.None)}
            dialogContentProps={{
              type: FluentDialogType.close,
              title: 'Filter Questions'
            }}
            modalProps={{
              isBlocking: true,
              styles: { main: { maxWidth: 450 } }
            }}
          >
            <FluentDropdown
              placeholder="Select filter"
              selectedKey={filterOption}
              onChange={(_, option) => setFilterOption(option?.key as string || 'all')}
              options={filterOptions}
              styles={{ dropdown: { width: '100%', marginBottom: '20px' } }}
            />
            <DialogFooter>
              <PrimaryButton onClick={() => setDialogType(DialogTypeEnum.None)} text="Apply" />
              <DefaultButton onClick={() => setDialogType(DialogTypeEnum.None)} text="Cancel" />
            </DialogFooter>
          </FluentDialog>
        );
      
      case DialogTypeEnum.Sort:
        return (
          <FluentDialog
            hidden={false}
            onDismiss={() => setDialogType(DialogTypeEnum.None)}
            dialogContentProps={{
              type: FluentDialogType.close,
              title: 'Sort Questions'
            }}
            modalProps={{
              isBlocking: true,
              styles: { main: { maxWidth: 450 } }
            }}
          >
            <FluentDropdown
              placeholder="Select sorting method"
              selectedKey={sortOption}
              onChange={(_, option) => setSortOption(option?.key as string || '')}
              options={sortOptions}
              styles={{ dropdown: { width: '100%', marginBottom: '20px' } }}
            />
            <DialogFooter>
              <PrimaryButton onClick={() => setDialogType(DialogTypeEnum.None)} text="Apply" />
              <DefaultButton onClick={() => setDialogType(DialogTypeEnum.None)} text="Cancel" />
            </DialogFooter>
          </FluentDialog>
        );
      
      default:
        return null;
    }
  };
  
  return (
    <div className={containerClass}>
      <div className={headerClass}>
        <FluentText variant="large" className={titleClass}>Quiz Questions</FluentText>
        <Stack horizontal tokens={{ childrenGap: 8 }} className={buttonGroupClass}>
          <PrimaryButton
            iconProps={addIcon}
            text="Add Question"
            onClick={handleAddQuestion}
          />
          <DefaultButton
            iconProps={importIcon}
            text="Import"
            onClick={handleImportQuestions}
          />
          <IconButton
            iconProps={shuffleIcon}
            title="Shuffle Questions"
            ariaLabel="Shuffle Questions"
            onClick={handleRandomize}
          />
          <IconButton
            iconProps={filterIcon}
            title="Filter Questions"
            ariaLabel="Filter Questions"
            onClick={() => setDialogType(DialogTypeEnum.Filter)}
          />
          <IconButton
            iconProps={sortIcon}
            title="Sort Questions"
            ariaLabel="Sort Questions"
            onClick={() => setDialogType(DialogTypeEnum.Sort)}
          />
          <IconButton
            iconProps={deleteIcon}
            title="Delete All Questions"
            ariaLabel="Delete All Questions"
            disabled={questions.length === 0}
            onClick={() => setDialogType(DialogTypeEnum.BulkDelete)}
          />
        </Stack>
      </div>
      
      {successMessage && (
        <FluentMessageBar
          messageBarType={MessageBarType.success}
          isMultiline={false}
          className={messageBarClass}
        >
          {successMessage}
        </FluentMessageBar>
      )}
      
      {questions.length === 0 ? (
        <div className={noQuestionsClass}>
          <FluentText>
            No questions have been added yet. Click &apos;Add Question&apos; to create your first question or &apos;Import&apos; to import questions.
          </FluentText>
        </div>
      ) : (
        <QuestionManagement 
          questions={filteredQuestions}
          onUpdateQuestions={onUpdateQuestions}
          onAddQuestion={handleAddQuestion}
          onEditQuestion={handleEditQuestion}
          onPreviewQuestion={handlePreviewQuestion}
          onDeleteQuestion={handleSingleDeleteQuestion}
        />
      )}
      
      {renderActiveDialog()}
    </div>
  );
};

export default QuizPropertyPane;