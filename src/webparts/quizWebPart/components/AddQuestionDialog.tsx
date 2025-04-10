import * as React from 'react';
import { useState, useEffect } from 'react';
import { IAddQuestionFormProps, IChoice, QuestionType, IQuizQuestion } from './interfaces';
import { v4 as uuidv4 } from 'uuid';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Checkbox,
  ChoiceGroup,
  IChoiceGroupOption,
  Stack,
  IStackTokens,
  IconButton,
  IIconProps,
  Text,
  Pivot,
  PivotItem,
  Toggle
} from '@fluentui/react';
import styles from './Quiz.module.scss';

// Icons for buttons
const addIcon: IIconProps = { iconName: 'Add' };
const saveIcon: IIconProps = { iconName: 'Save' };
const deleteIcon: IIconProps = { iconName: 'Delete' };
const previewIcon: IIconProps = { iconName: 'RedEye' };
const resetIcon: IIconProps = { iconName: 'Refresh' };
const cancelIcon: IIconProps = { iconName: 'Cancel' };

// Stack token for spacing
const stackTokens: IStackTokens = {
  childrenGap: 15
};

const AddQuestionDialog: React.FC<IAddQuestionFormProps> = ({
  categories,
  onSubmit,
  onCancel,
  isSubmitting,
  onPreviewQuestion,
  initialQuestion
}) => {
  const [title, setTitle] = useState('');
  const [category, setCategory] = useState('');
  const [newCategory, setNewCategory] = useState('');
  const [questionType, setQuestionType] = useState(QuestionType.MultipleChoice);
  const [choices, setChoices] = useState<IChoice[]>([
    { id: uuidv4(), text: '', isCorrect: false },
    { id: uuidv4(), text: '', isCorrect: false },
  ]);
  const [correctChoiceId, setCorrectChoiceId] = useState('');
  const [shortAnswerText, setShortAnswerText] = useState('');
  const [explanation, setExplanation] = useState('');
  const [points, setPoints] = useState('1');
  const [validationError, setValidationError] = useState('');
  const [activeTab, setActiveTab] = useState('question');
  const [caseSensitive, setCaseSensitive] = useState(false); // For short answer questions
  
  // Initialize form if editing an existing question
  useEffect(() => {
    if (initialQuestion) {
      setTitle(initialQuestion.title);
      setCategory(initialQuestion.category);
      setQuestionType(initialQuestion.type);
      setChoices([...initialQuestion.choices]);
      
      // Set correct choice ID for multiple choice
      if (initialQuestion.type === QuestionType.MultipleChoice || 
          initialQuestion.type === QuestionType.TrueFalse) {
        const correct = initialQuestion.choices.find(c => c.isCorrect);
        if (correct) {
          setCorrectChoiceId(correct.id);
        }
      }
      
      // Set short answer text
      if (initialQuestion.type === QuestionType.ShortAnswer && initialQuestion.correctAnswer) {
        setShortAnswerText(initialQuestion.correctAnswer);
        // If case sensitivity setting exists
        if (initialQuestion.caseSensitive !== undefined) {
          setCaseSensitive(initialQuestion.caseSensitive);
        }
      }
      
      // Set explanation and points
      if (initialQuestion.explanation) {
        setExplanation(initialQuestion.explanation);
      }
      
      if (initialQuestion.points) {
        setPoints(initialQuestion.points.toString());
      }
    }
  }, [initialQuestion]);

  const handleChoiceTextChange = (id: string, text: string): void => {
    setChoices(choices.map((c: IChoice) => (c.id === id ? { ...c, text } : c)));
  };

  const handleChoiceCorrectChange = (id: string, isCorrect: boolean): void => {
    // For multiple choice, only one answer can be correct
    if (questionType === QuestionType.MultipleChoice && isCorrect) {
      setChoices(choices.map(c => ({ ...c, isCorrect: c.id === id })));
      setCorrectChoiceId(id);
    } else {
      // For multi-select, multiple answers can be correct
      setChoices(choices.map(c => (c.id === id ? { ...c, isCorrect } : c)));
    }
  };

  const handleSubmit = (): void => {
    // Validation
    if (!title.trim()) {
      setValidationError('Question title is required.');
      return;
    }
    
    if (!category && !newCategory) {
      setValidationError('Please select or enter a category.');
      return;
    }
    
    // Validate based on question type
    if (questionType === QuestionType.MultipleChoice || questionType === QuestionType.TrueFalse) {
      if (choices.filter(c => c.text.trim()).length < 2) {
        setValidationError('At least 2 valid choices are required.');
        return;
      }
      
      if (!choices.some(c => c.isCorrect)) {
        setValidationError('Please mark at least one choice as correct.');
        return;
      }
    } else if (questionType === QuestionType.MultiSelect) {
      if (choices.filter(c => c.text.trim()).length < 2) {
        setValidationError('At least 2 valid choices are required.');
        return;
      }
      
      if (!choices.some(c => c.isCorrect)) {
        setValidationError('Please mark at least one choice as correct.');
        return;
      }
    } else if (questionType === QuestionType.ShortAnswer) {
      if (!shortAnswerText.trim()) {
        setValidationError('Please enter the correct answer for the short answer question.');
        return;
      }
    }

    // Points validation
    const pointsValue = parseInt(points, 10);
    if (isNaN(pointsValue) || pointsValue < 1) {
      setValidationError('Points must be a positive number.');
      return;
    }

    // Create question object
    const newQuestion: IQuizQuestion = {
      id: initialQuestion ? initialQuestion.id : Date.now(),
      title,
      category: category === 'new' ? newCategory : category,
      type: questionType,
      choices: choices.filter(c => c.text.trim()), // Filter out empty choices
      correctAnswer: questionType === QuestionType.ShortAnswer ? shortAnswerText : undefined,
      explanation: explanation.trim() || undefined,
      points: pointsValue,
      caseSensitive: questionType === QuestionType.ShortAnswer ? caseSensitive : undefined,
      // Add timestamp for tracking
      lastModified: new Date().toISOString()
    };

    onSubmit(newQuestion);
  };

  const handlePreview = (): void => {
    // Create question object for preview
    const previewQuestion: IQuizQuestion = {
      id: initialQuestion ? initialQuestion.id : Date.now(),
      title,
      category: category === 'new' ? newCategory : category,
      type: questionType,
      choices: choices.filter(c => c.text.trim()), // Filter out empty choices
      correctAnswer: questionType === QuestionType.ShortAnswer ? shortAnswerText : undefined,
      explanation: explanation.trim() || undefined,
      points: parseInt(points, 10) || 1,
      caseSensitive: questionType === QuestionType.ShortAnswer ? caseSensitive : undefined
    };

    onPreviewQuestion(previewQuestion);
  };

  const resetForm = (): void => {
    setTitle('');
    setCategory('');
    setNewCategory('');
    setQuestionType(QuestionType.MultipleChoice);
    setChoices([
      { id: uuidv4(), text: '', isCorrect: false },
      { id: uuidv4(), text: '', isCorrect: false },
    ]);
    setCorrectChoiceId('');
    setShortAnswerText('');
    setExplanation('');
    setPoints('1');
    setCaseSensitive(false);
    setValidationError('');
  };

  // Generate category dropdown options
  const categoryOptions: IDropdownOption[] = [
    ...categories.map(cat => ({ key: cat, text: cat })),
    { key: 'new', text: 'Add new category' }
  ];

  // Generate question type options
  const questionTypeOptions: IDropdownOption[] = [
    { key: QuestionType.MultipleChoice, text: 'Multiple Choice' },
    { key: QuestionType.TrueFalse, text: 'True/False' },
    { key: QuestionType.MultiSelect, text: 'Multiple Select' },
    { key: QuestionType.ShortAnswer, text: 'Short Answer' }
  ];

  // For True/False, create choice group options
  const tfOptions: IChoiceGroupOption[] = [
    { key: 'true', text: 'True' },
    { key: 'false', text: 'False' }
  ];

  // Render different inputs based on question type
  const renderQuestionTypeInputs = (): JSX.Element | null => {
    switch (questionType) {
      case QuestionType.MultipleChoice:
        return (
          <div className={styles.choicesContainer}>
            <Text className={styles.formSectionTitle}>Choices (select the correct answer)</Text>
            {choices.map((choice, idx) => (
              <Stack horizontal tokens={stackTokens} verticalAlign="center" key={choice.id} className={styles.choiceRow}>
                <Checkbox
                  checked={choice.isCorrect}
                  onChange={(_e, checked) => handleChoiceCorrectChange(choice.id, !!checked)}
                  label=""
                  styles={{ root: { marginRight: 8 } }}
                />
                <TextField
                  placeholder={`Choice ${idx + 1}`}
                  value={choice.text}
                  onChange={(_e, value) => handleChoiceTextChange(choice.id, value || '')}
                  styles={{ root: { flexGrow: 1 } }}
                />
                <IconButton
                  iconProps={deleteIcon}
                  title="Delete"
                  ariaLabel="Delete"
                  onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                  disabled={choices.length <= 2} // Require at least 2 choices
                  className={styles.actionButton}
                />
              </Stack>
            ))}

            <Stack horizontal className={styles.buttonGroup}>
              <DefaultButton
                iconProps={addIcon}
                text="Add Choice"
                onClick={() =>
                  setChoices([...choices, { id: uuidv4(), text: '', isCorrect: false }])
                }
              />
            </Stack>
          </div>
        );
        
      case QuestionType.TrueFalse:
        return (
          <div className={styles.choicesContainer}>
            <Text className={styles.formSectionTitle}>Select the correct answer:</Text>
            <ChoiceGroup
              options={tfOptions}
              selectedKey={correctChoiceId}
              onChange={(_e, option) => {
                if (option) {
                  setCorrectChoiceId(option.key);
                  // Update the choices array to mark the correct answer
                  setChoices([
                    { id: 'true', text: 'True', isCorrect: option.key === 'true' },
                    { id: 'false', text: 'False', isCorrect: option.key === 'false' }
                  ]);
                }
              }}
            />
          </div>
        );
        
      case QuestionType.MultiSelect: {
        return (
          <div className={styles.choicesContainer}>
            <Text className={styles.formSectionTitle}>Choices (select all correct answers)</Text>
            {choices.map((choice, idx) => (
              <Stack horizontal tokens={stackTokens} verticalAlign="center" key={choice.id} className={styles.choiceRow}>
                <Checkbox
                  checked={choice.isCorrect}
                  onChange={(_e, checked) => handleChoiceCorrectChange(choice.id, !!checked)}
                  label=""
                  styles={{ root: { marginRight: 8 } }}
                />
                <TextField
                  value={choice.text}
                  onChange={(_e, value) => handleChoiceTextChange(choice.id, value || '')}
                  styles={{ root: { flexGrow: 1 } }}
                  placeholder={`Choice ${idx + 1}`}
                />
                <IconButton
                  iconProps={deleteIcon}
                  title="Delete"
                  ariaLabel="Delete"
                  onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                  disabled={choices.length <= 2} // Require at least 2 choices
                  className={styles.actionButton}
                />
              </Stack>
            ))}

            <Stack horizontal className={styles.buttonGroup}>
              <DefaultButton
                iconProps={addIcon}
                text="Add Choice"
                onClick={() =>
                  setChoices([...choices, { id: uuidv4(), text: '', isCorrect: false }])
                }
              />
            </Stack>
          </div>
        );
      }
        
      case QuestionType.ShortAnswer:
        return (
          <div className={styles.choicesContainer}>
            <Text className={styles.formSectionTitle}>Correct Answer</Text>
            <TextField
              required
              value={shortAnswerText}
              onChange={(_e, value) => setShortAnswerText(value || '')}
              placeholder="Enter the correct answer"
            />
            <Toggle
              label="Case sensitive"
              checked={caseSensitive}
              onChange={(_e, checked) => setCaseSensitive(!!checked)}
              onText="On"
              offText="Off"
              inlineLabel
            />
          </div>
        );
        
      default:
        return null;
    }
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={onCancel}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: initialQuestion ? 'Edit Question' : 'Add New Question'
      }}
      modalProps={{
        isBlocking: true,
        styles: { main: { maxWidth: '800px', minWidth: '600px' } }
      }}
    >
      {validationError && (
        <MessageBar 
          messageBarType={MessageBarType.error}
          isMultiline={false} 
          dismissButtonAriaLabel="Close"
          styles={{ root: { marginBottom: 15 } }}
        >
          {validationError}
        </MessageBar>
      )}

      <Pivot 
        selectedKey={activeTab} 
        onLinkClick={(item) => item && setActiveTab(item.props.itemKey || 'question')}
        styles={{ root: { marginBottom: 20 } }}
      >
        <PivotItem headerText="Question" itemKey="question" />
        <PivotItem headerText="Additional Info" itemKey="additional" />
      </Pivot>

      {activeTab === 'question' && (
        <Stack tokens={stackTokens} className={styles.formWrapper}>
          <TextField
            label="Question"
            required
            value={title}
            onChange={(_e, value) => setTitle(value || '')}
            placeholder="Enter your question here"
            styles={{ fieldGroup: { width: '100%' } }}
          />

          <Dropdown
            label="Question Type"
            required
            selectedKey={questionType}
            onChange={(_e, option) => {
              if (option) {
                const newType = option.key as QuestionType;
                setQuestionType(newType);
                
                // Reset choices based on type
                if (newType === QuestionType.TrueFalse) {
                  setChoices([
                    { id: 'true', text: 'True', isCorrect: false },
                    { id: 'false', text: 'False', isCorrect: false }
                  ]);
                  setCorrectChoiceId('');
                } else if (newType === QuestionType.MultipleChoice || newType === QuestionType.MultiSelect) {
                  // Keep existing choices or reset if needed
                  if (choices.length < 2) {
                    setChoices([
                      { id: uuidv4(), text: '', isCorrect: false },
                      { id: uuidv4(), text: '', isCorrect: false }
                    ]);
                  }
                }
              }
            }}
            options={questionTypeOptions}
          />

          <Dropdown
            label="Category"
            required
            selectedKey={category}
            onChange={(_e, option) => option && setCategory(option.key as string)}
            options={categoryOptions}
            placeholder="Select or add category"
          />

          {category === 'new' && (
            <TextField
              label="New Category"
              required
              value={newCategory}
              onChange={(_e, value) => setNewCategory(value || '')}
              placeholder="Enter new category"
              styles={{ fieldGroup: { width: '100%' } }}
            />
          )}

          {renderQuestionTypeInputs()}
        </Stack>
      )}

      {activeTab === 'additional' && (
        <Stack tokens={stackTokens} className={styles.formWrapper}>
          <TextField
            label="Points"
            type="number"
            value={points}
            onChange={(_e, value) => {
              // Ensure only positive numbers are entered
              const numericValue = value ? value.replace(/[^0-9]/g, '') : '';
              setPoints(numericValue || '1');
            }}
            min="1"
            placeholder="1"
            styles={{ fieldGroup: { width: '100px' } }}
          />

          <TextField
            label="Explanation (Optional)"
            multiline
            rows={4}
            value={explanation}
            onChange={(_e, value) => setExplanation(value || '')}
            placeholder="Enter an explanation for the correct answer"
            styles={{ fieldGroup: { width: '100%' } }}
          />
        </Stack>
      )}

      <DialogFooter className={styles.formButtons}>
        <PrimaryButton
          onClick={handleSubmit}
          text={initialQuestion ? 'Update Question' : 'Save Question'}
          disabled={isSubmitting}
          iconProps={saveIcon}
        />
        <DefaultButton
          onClick={handlePreview}
          text="Preview"
          iconProps={previewIcon}
        />
        {!initialQuestion && (
          <DefaultButton
            onClick={resetForm}
            text="Reset"
            iconProps={resetIcon}
          />
        )}
        <DefaultButton
          onClick={onCancel}
          text="Cancel"
          disabled={isSubmitting}
          iconProps={cancelIcon}
        />
      </DialogFooter>
    </Dialog>
  );
};

export default AddQuestionDialog;