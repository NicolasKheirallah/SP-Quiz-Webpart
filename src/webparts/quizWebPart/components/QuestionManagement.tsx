import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  CommandBar,
  ICommandBarItemProps,
  SearchBox,
  Dropdown,
  IDropdownOption,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  IStackTokens
} from '@fluentui/react';
import { IQuizQuestion, QuestionType } from './interfaces';
import styles from './Quiz.module.scss';

// Define interface for component props
export interface IQuestionManagementProps {
  questions: IQuizQuestion[];
  onUpdateQuestions: (questions: IQuizQuestion[]) => void;
  onAddQuestion: () => void;
  onEditQuestion: (question: IQuizQuestion) => void;
  onPreviewQuestion: (question: IQuizQuestion) => void;
  onDeleteQuestion?: (questionId: number) => void;
}

// Stack tokens for spacing
const stackTokens: IStackTokens = {
  childrenGap: 8
};

// Component implementation
const QuestionManagement: React.FC<IQuestionManagementProps> = (props) => {
  const { questions, onUpdateQuestions, onAddQuestion, onEditQuestion, onPreviewQuestion, onDeleteQuestion } = props;

  // State
  const [filteredQuestions, setFilteredQuestions] = useState<IQuizQuestion[]>(questions);
  const [searchText, setSearchText] = useState('');
  const [categoryFilter, setCategoryFilter] = useState<string>('');
  const [typeFilter, setTypeFilter] = useState<string>('');
  const [selectedItems, setSelectedItems] = useState<IQuizQuestion[]>([]);
  const [isDeleteDialogVisible, setIsDeleteDialogVisible] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<'selected' | 'single' | null>(null);
  const [questionToDelete, setQuestionToDelete] = useState<IQuizQuestion | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [successMessage, setSuccessMessage] = useState('');

  // Create a selection instance
  const [selection] = useState(() => new Selection({
    onSelectionChanged: () => {
      setSelectedItems(selection.getSelection() as IQuizQuestion[]);
    }
  }));

  // Format question type for display
  const formatQuestionType = useCallback((type: QuestionType): string => {
    switch (type) {
      case QuestionType.MultipleChoice:
        return 'Multiple Choice';
      case QuestionType.TrueFalse:
        return 'True/False';
      case QuestionType.MultiSelect:
        return 'Multiple Select';
      case QuestionType.ShortAnswer:
        return 'Short Answer';
      case QuestionType.Matching:
        return 'Matching';
      default:
        return String(type);
    }
  }, []);
  
  // Apply filters to questions
  const applyFilters = useCallback(() => {
    let filtered = [...questions];

    // Apply search filter
    if (searchText) {
      const search = searchText.toLowerCase();
      filtered = filtered.filter(
        (q) => q.title.toLowerCase().includes(search) || q.category.toLowerCase().includes(search)
      );
    }

    // Apply category filter
    if (categoryFilter) {
      filtered = filtered.filter((q) => q.category === categoryFilter);
    }

    // Apply type filter
    if (typeFilter) {
      filtered = filtered.filter((q) => q.type === typeFilter);
    }

    setFilteredQuestions(filtered);
  }, [questions, searchText, categoryFilter, typeFilter]);

  // Handle single question delete click
  const handleSingleDeleteClick = useCallback((question: IQuizQuestion): void => {
    if (onDeleteQuestion) {
      // If parent provided a delete handler, use it directly
      onDeleteQuestion(question.id);
      setSuccessMessage(`Question "${question.title}" has been deleted.`);
    } else {
      // Otherwise show confirmation dialog
      setQuestionToDelete(question);
      setDeleteTarget('single');
      setIsDeleteDialogVisible(true);
    }
  }, [onDeleteQuestion]);

  // Handle bulk delete click
  const handleBulkDeleteClick = useCallback((): void => {
    setDeleteTarget('selected');
    setIsDeleteDialogVisible(true);
  }, []);

  // Handle search text change
  const handleSearchChange = useCallback((event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    setSearchText(newValue || '');
  }, []);

  // Handle category filter change
  const handleCategoryChange = useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    setCategoryFilter(option?.key as string || '');
  }, []);

  // Handle type filter change
  const handleTypeChange = useCallback((event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    setTypeFilter(option?.key as string || '');
  }, []);

  // Handle delete confirmation
  const handleDeleteConfirm = useCallback((): void => {
    setIsLoading(true);
    setErrorMessage('');

    try {
      let updatedQuestions: IQuizQuestion[] = [...questions];

      if (deleteTarget === 'single' && questionToDelete) {
        updatedQuestions = questions.filter((q) => q.id !== questionToDelete.id);
        setSuccessMessage(`Question "${questionToDelete.title}" has been deleted.`);
      } else if (deleteTarget === 'selected' && selectedItems.length > 0) {
        const selectedIds = new Set(selectedItems.map((item) => item.id));
        updatedQuestions = questions.filter((q) => !selectedIds.has(q.id));
        setSuccessMessage(`${selectedItems.length} questions have been deleted.`);
      }

      // Update questions through the callback
      onUpdateQuestions(updatedQuestions);
      
      // Close dialog and reset state
      setIsDeleteDialogVisible(false);
      setQuestionToDelete(null);
      setDeleteTarget(null);
      
      // Reset selection
      selection.setAllSelected(false);
      
      // Clear success message after 3 seconds
      setTimeout(() => {
        setSuccessMessage('');
      }, 3000);
    } catch (error) {
      setErrorMessage('An error occurred while deleting questions.');
      console.error('Delete error:', error);
    } finally {
      setIsLoading(false);
    }
  }, [questions, deleteTarget, questionToDelete, selectedItems, onUpdateQuestions, selection]);

  // Handle delete cancellation
  const handleDeleteCancel = useCallback((): void => {
    setIsDeleteDialogVisible(false);
    setQuestionToDelete(null);
    setDeleteTarget(null);
  }, []);

  // Delete dialog content
  const getDeleteDialogContent = useCallback((): string => {
    if (deleteTarget === 'single' && questionToDelete) {
      return `Are you sure you want to delete the question "${questionToDelete.title}"?`;
    } else if (deleteTarget === 'selected') {
      return `Are you sure you want to delete ${selectedItems.length} selected question(s)?`;
    }
    return '';
  }, [deleteTarget, questionToDelete, selectedItems]);

  // Extract unique categories and types for filters
  const categories: IDropdownOption[] = React.useMemo(() => {
    const uniqueCategories = Array.from(new Set(questions.map((q) => q.category))).filter(Boolean);
    return [
      { key: '', text: 'All Categories' },
      ...uniqueCategories.map((category) => ({ key: category, text: category })),
    ];
  }, [questions]);

  const types: IDropdownOption[] = React.useMemo(() => {
    const uniqueTypes = Array.from(new Set(questions.map((q) => q.type)));
    return [
      { key: '', text: 'All Types' },
      ...uniqueTypes.map((type) => ({ key: type, text: formatQuestionType(type) })),
    ];
  }, [questions, formatQuestionType]);

  // Command bar items
  const getCommandBarItems = useCallback((): ICommandBarItemProps[] => {
    const items: ICommandBarItemProps[] = [
      {
        key: 'addQuestion',
        text: 'Add Question',
        iconProps: { iconName: 'Add' },
        onClick: onAddQuestion,
      },
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: () => applyFilters(),
      },
    ];

    if (selectedItems.length > 0) {
      items.push({
        key: 'delete',
        text: 'Delete Selected',
        iconProps: { iconName: 'Delete' },
        onClick: handleBulkDeleteClick,
      });
    }

    return items;
  }, [onAddQuestion, applyFilters, selectedItems.length, handleBulkDeleteClick]);

  // Columns for DetailsList
  const columns: IColumn[] = React.useMemo(() => [
    {
      key: 'title',
      name: 'Question',
      fieldName: 'title',
      minWidth: 200,
      maxWidth: 400,
      isResizable: true,
      data: 'string',
      isPadded: true,
    },
    {
      key: 'category',
      name: 'Category',
      fieldName: 'category',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      data: 'string',
      isPadded: true,
    },
    {
      key: 'type',
      name: 'Type',
      fieldName: 'type',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      data: 'string',
      onRender: (item: IQuizQuestion) => {
        return <span>{formatQuestionType(item.type)}</span>;
      },
      isPadded: true,
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 100,
      maxWidth: 100,
      isResizable: false,
      onRender: (item: IQuizQuestion) => {
        return (
          <Stack horizontal tokens={stackTokens} className={styles.actionButtons}>
            <IconButton
              iconProps={{ iconName: 'View' }}
              title="Preview"
              ariaLabel="Preview"
              onClick={() => onPreviewQuestion(item)}
              className={styles.actionButton}
            />
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              ariaLabel="Edit"
              onClick={() => onEditQuestion(item)}
              className={styles.actionButton}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete"
              onClick={() => handleSingleDeleteClick(item)}
              className={styles.actionButton}
            />
          </Stack>
        );
      },
    },
  ], [formatQuestionType, handleSingleDeleteClick, onEditQuestion, onPreviewQuestion]);

  // Update filtered questions when questions array changes or filters change
  useEffect(() => {
    applyFilters();
    // Reset selection when questions change
    if (selection) {
      selection.setAllSelected(false);
    }
  }, [questions, searchText, categoryFilter, typeFilter, applyFilters, selection]);

  // Clear success message after 3 seconds
  useEffect(() => {
    if (successMessage) {
      const timer = setTimeout(() => {
        setSuccessMessage('');
      }, 3000);
      return () => clearTimeout(timer);
    }
  }, [successMessage]);

  return (
    <div>
      {errorMessage && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          dismissButtonAriaLabel="Close"
          className={styles.statusBar}
        >
          {errorMessage}
        </MessageBar>
      )}

      {successMessage && (
        <MessageBar
          messageBarType={MessageBarType.success}
          isMultiline={false}
          dismissButtonAriaLabel="Close"
          className={`${styles.statusBar} ${styles.success}`}
        >
          {successMessage}
        </MessageBar>
      )}

      <div className={styles.commandBar}>
        <CommandBar items={getCommandBarItems()} />
      </div>

      <div className={styles.filtersBar}>
        <SearchBox
          placeholder="Search questions..."
          onChange={handleSearchChange}
          value={searchText}
          className={styles.searchBox}
        />
        <Dropdown
          placeholder="Filter by category"
          options={categories}
          selectedKey={categoryFilter}
          onChange={handleCategoryChange}
          className={styles.filterDropdown}
        />
        <Dropdown
          placeholder="Filter by type"
          options={types}
          selectedKey={typeFilter}
          onChange={handleTypeChange}
          className={styles.filterDropdown}
        />
      </div>

      {selectedItems.length > 0 && (
        <div className={`${styles.statusBar} ${styles.info}`}>
          <Text>{selectedItems.length} item(s) selected</Text>
        </div>
      )}

      <DetailsList
        items={filteredQuestions}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selection={selection}
        selectionMode={SelectionMode.multiple}
        selectionPreservedOnEmptyClick={true}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="select row"
      />

      {filteredQuestions.length === 0 && !isLoading && (
        <Stack horizontalAlign="center" verticalAlign="center" className={styles.emptyState}>
          <Text variant="medium">No questions found. Try different search criteria or add new questions.</Text>
        </Stack>
      )}

      <Dialog
        hidden={!isDeleteDialogVisible}
        onDismiss={handleDeleteCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Confirm Delete',
          subText: getDeleteDialogContent(),
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 450 } },
        }}
      >
        {isLoading && <Spinner size={SpinnerSize.small} label="Deleting..." />}
        <DialogFooter>
          <PrimaryButton onClick={handleDeleteConfirm} text="Delete" disabled={isLoading} />
          <DefaultButton onClick={handleDeleteCancel} text="Cancel" disabled={isLoading} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default QuestionManagement;