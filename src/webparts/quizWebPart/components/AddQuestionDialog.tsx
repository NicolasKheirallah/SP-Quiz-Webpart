import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  IAddQuestionFormProps,
  IChoice,
  QuestionType,
  IQuizQuestion,
  IMatchingPair,
  IQuizImage,
  ICodeSnippet
} from './interfaces';
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
  Toggle,
  IDialogContentProps,
  IModalProps,
  ITextStyles,
  SpinButton,
  ISpinButtonStyles,
  mergeStyleSets,
  IProcessedStyleSet
} from '@fluentui/react';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import styles from './Quiz.module.scss';
import ImageUpload from './ImageUpload';
import CodeSnippet from './CodeSnippet';

// Icons for buttons
const addIcon: IIconProps = { iconName: 'Add' };
const saveIcon: IIconProps = { iconName: 'Save' };
const deleteIcon: IIconProps = { iconName: 'Delete' };
const previewIcon: IIconProps = { iconName: 'RedEye' };
const resetIcon: IIconProps = { iconName: 'Refresh' };
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const imageIcon: IIconProps = { iconName: 'Picture' };
const codeIcon: IIconProps = { iconName: 'Code' };


// Stack tokens for choice row
const choiceRowStackTokens: IStackTokens = {
  childrenGap: 12
};

// Custom styles for the form section title
const formSectionTitleStyles: ITextStyles = {
  root: {
    fontSize: '16px',
    fontWeight: 600,
    marginBottom: '12px',
    color: '#323130',
    marginTop: '8px'
  }
};

// Spin button styles
const spinButtonStyles: Partial<ISpinButtonStyles> = {
  spinButtonWrapper: {
    width: 120
  }
};

// Custom styles for the dialog components
const customStyles: IProcessedStyleSet<{
  dialogRoot: string;
  dialogContent: string;
  formContainer: string;
  sectionContainer: string;
  fieldGroup: string;
  choicesContainer: string;
  matchingContainer: string;
  horizontalGroup: string;
  footer: string;
}> = mergeStyleSets({
  dialogRoot: {
    selectors: {
      '@media (min-width: 480px)': {
        minWidth: '800px !important',
        maxWidth: '90vw !important'
      }
    }
  },
  dialogContent: {
    padding: '0 24px 20px 24px',
    selectors: {
      '@media (max-width: 480px)': {
        padding: '0 16px 16px 16px'
      }
    }
  },
  formContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
    padding: '4px 4px 16px 4px',
    overflowY: 'visible'
  },
  sectionContainer: {
    marginBottom: '20px',
    padding: '16px',
    backgroundColor: '#fafafa',
    border: '1px solid #edebe9',
    borderRadius: '4px'
  },
  fieldGroup: {
    marginBottom: '16px'
  },
  choicesContainer: {
    padding: '20px',
    backgroundColor: '#f8f8f8',
    borderRadius: '4px',
    border: '1px solid #edebe9',
    marginTop: '16px',
    marginBottom: '16px'
  },
  matchingContainer: {
    padding: '20px',
    backgroundColor: '#f8f8f8',
    borderRadius: '4px',
    border: '1px solid #edebe9',
    marginTop: '16px',
    marginBottom: '16px'
  },
  horizontalGroup: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '16px',
    selectors: {
      '@media (max-width: 640px)': {
        flexDirection: 'column'
      }
    }
  },
  footer: {
    marginTop: '24px',
    padding: '16px 0',
    borderTop: '1px solid #edebe9'
  }
});

  // Custom modal props with fixed width issue and single scrollbar
const customModalProps: IModalProps = {
  isBlocking: true,
  styles: {
    main: {
      selectors: {
        '@media (min-width: 480px)': {
          minWidth: '800px !important',
          maxWidth: '90vw !important',
          width: 'auto !important'
        }
      }
    },
    scrollableContent: {
      maxHeight: '90vh'
    }
  },
  className: 'wideFormDialog'
};

const AddQuestionDialog: React.FC<IAddQuestionFormProps> = ({
  categories,
  onSubmit,
  onCancel,
  isSubmitting,
  onPreviewQuestion,
  initialQuestion,
  context
}) => {
  const [title, setTitle] = useState('');
  const [category, setCategory] = useState('');
  const [newCategory, setNewCategory] = useState('');
  const [questionType, setQuestionType] = useState(QuestionType.MultipleChoice);
  const [choices, setChoices] = useState<IChoice[]>([
    { id: uuidv4(), text: '', isCorrect: false },
    { id: uuidv4(), text: '', isCorrect: false },
  ]);
  const [matchingPairs, setMatchingPairs] = useState<IMatchingPair[]>([
    { id: uuidv4(), leftItem: '', rightItem: '' },
    { id: uuidv4(), leftItem: '', rightItem: '' }
  ]);

  const [correctChoiceId, setCorrectChoiceId] = useState('');
  const [shortAnswerText, setShortAnswerText] = useState('');
  const [explanation, setExplanation] = useState('');
  const [points, setPoints] = useState('1');
  const [validationError, setValidationError] = useState('');
  const [activeTab, setActiveTab] = useState('question');
  const [caseSensitive, setCaseSensitive] = useState(false); // For short answer questions

  // New fields for extended features
  const [images, setImages] = useState<IQuizImage[]>([]);
  const [codeSnippets, setCodeSnippets] = useState<ICodeSnippet[]>([]);
  const [timeLimit, setTimeLimit] = useState<number | undefined>(undefined);
  const [addingImage, setAddingImage] = useState(false);
  const [addingCodeSnippet, setAddingCodeSnippet] = useState(false);
  const [timeLimitEnabled, setTimeLimitEnabled] = useState<boolean>(!!initialQuestion?.timeLimit);

  // For rich text description
  const [description, setDescription] = useState<string>('');

  // Initialize form if editing an existing question
  useEffect(() => {
    if (initialQuestion) {
      setTitle(initialQuestion.title);
      setCategory(initialQuestion.category);
      setQuestionType(initialQuestion.type);
      setChoices([...initialQuestion.choices]);

      // Set description if it exists
      if (initialQuestion.description) {
        setDescription(initialQuestion.description);
      }

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

      // Set new feature fields
      if (initialQuestion.images && initialQuestion.images.length > 0) {
        setImages([...initialQuestion.images]);
      }

      if (initialQuestion.codeSnippets && initialQuestion.codeSnippets.length > 0) {
        setCodeSnippets([...initialQuestion.codeSnippets]);
      }

      if (initialQuestion.timeLimit) {
        setTimeLimit(initialQuestion.timeLimit);
      }
      if (initialQuestion.type === QuestionType.Matching && initialQuestion.matchingPairs) {
        setMatchingPairs([...initialQuestion.matchingPairs]);
      }
    }
  }, [initialQuestion]);

  // Function to handle rich text changes - must return the string
  const handleDescriptionChange = (newText: string): string => {
    setDescription(newText);
    return newText;
  };

  // Handle image upload
  const handleImageUpload = (image: IQuizImage): void => {
    // Check if we already have this image (updating)
    const existingIndex = images.findIndex(img => img.id === image.id);

    if (existingIndex >= 0) {
      // Update existing image
      const updatedImages = [...images];
      updatedImages[existingIndex] = image;
      setImages(updatedImages);
    } else {
      // Add new image
      setImages([...images, image]);
    }

    setAddingImage(false);
  };

  // Handle image removal
  const handleImageRemove = (imageId: string): void => {
    setImages(images.filter(img => img.id !== imageId));
  };

  // Handle code snippet updates
  const handleCodeSnippetChange = (snippet: ICodeSnippet): void => {
    // Check if we already have this snippet (updating)
    const existingIndex = codeSnippets.findIndex(s => s.id === snippet.id);

    if (existingIndex >= 0) {
      // Update existing snippet
      const updatedSnippets = [...codeSnippets];
      updatedSnippets[existingIndex] = snippet;
      setCodeSnippets(updatedSnippets);
    } else {
      // Add new snippet
      setCodeSnippets([...codeSnippets, snippet]);
    }

    setAddingCodeSnippet(false);
  };

  // Handle code snippet removal
  const handleCodeSnippetRemove = (snippetId: string): void => {
    setCodeSnippets(codeSnippets.filter(s => s.id !== snippetId));
  };

  // Handle time limit change
  const handleTimeLimitChange = (value: string): void => {
    const numValue = parseInt(value, 10);
    setTimeLimit(isNaN(numValue) ? undefined : numValue);
  };

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

  // Handle image assignment to a choice
  const handleChoiceImageUpload = (choiceId: string, image: IQuizImage): void => {
    setChoices(choices.map(c =>
      c.id === choiceId
        ? { ...c, image }
        : c
    ));
  };

  // Handle image removal from a choice
  const handleChoiceImageRemove = (choiceId: string): void => {
    setChoices(choices.map(c =>
      c.id === choiceId
        ? { ...c, image: undefined }
        : c
    ));
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
    } else if (questionType === QuestionType.Matching) {
      // Validate matching pairs
      if (matchingPairs.length < 2) {
        setValidationError('At least 2 matching pairs are required.');
        return;
      }
      
      // Check if all pairs have both left and right items
      if (matchingPairs.some(pair => !pair.leftItem.trim() || !pair.rightItem.trim())) {
        setValidationError('All matching pairs must have both left and right items.');
        return;
      }
    }

    // Points validation
    const pointsValue = parseInt(points, 10);
    if (isNaN(pointsValue) || pointsValue < 1) {
      setValidationError('Points must be a positive number.');
      return;
    }

    // Time limit validation (if set)
    if (timeLimit !== undefined && (isNaN(timeLimit) || timeLimit < 5 || timeLimit > 3600)) {
      setValidationError('Time limit must be between 5 seconds and 3600 seconds (1 hour).');
      return;
    }

    // Create question object
    const newQuestion: IQuizQuestion = {
      id: initialQuestion ? initialQuestion.id : Date.now(),
      title,
      description: description, // Add the rich text description
      category: category === 'new' ? newCategory : category,
      type: questionType,
      choices: choices.filter(c => c.text.trim()), // Filter out empty choices
      correctAnswer: questionType === QuestionType.ShortAnswer ? shortAnswerText : undefined,
      explanation: explanation.trim() || undefined,
      points: pointsValue,
      caseSensitive: questionType === QuestionType.ShortAnswer ? caseSensitive : undefined,
      // Add matching pairs if it's a matching question
      matchingPairs: questionType === QuestionType.Matching ? matchingPairs : undefined,
      // Add timestamp for tracking
      lastModified: new Date().toISOString(),
      // Add new feature fields
      images: images.length > 0 ? images : undefined,
      codeSnippets: codeSnippets.length > 0 ? codeSnippets : undefined,
      timeLimit: timeLimitEnabled ? timeLimit : undefined
    };

    onSubmit(newQuestion);
  };

  const handlePreview = (): void => {
    // Create question object for preview
    const previewQuestion: IQuizQuestion = {
      id: initialQuestion ? initialQuestion.id : Date.now(),
      title,
      description: description, // Add the rich text description
      category: category === 'new' ? newCategory : category,
      type: questionType,
      choices: choices.filter(c => c.text.trim()), // Filter out empty choices
      correctAnswer: questionType === QuestionType.ShortAnswer ? shortAnswerText : undefined,
      explanation: explanation.trim() || undefined,
      points: parseInt(points, 10) || 1,
      caseSensitive: questionType === QuestionType.ShortAnswer ? caseSensitive : undefined,
      // Add matching pairs if it's a matching question
      matchingPairs: questionType === QuestionType.Matching ? matchingPairs : undefined,
      // Add new feature fields for preview
      images: images.length > 0 ? images : undefined,
      codeSnippets: codeSnippets.length > 0 ? codeSnippets : undefined,
      timeLimit: timeLimit
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
    setMatchingPairs([
      { id: uuidv4(), leftItem: '', rightItem: '' },
      { id: uuidv4(), leftItem: '', rightItem: '' }
    ]);
    setCorrectChoiceId('');
    setShortAnswerText('');
    setExplanation('');
    setPoints('1');
    setCaseSensitive(false);
    setValidationError('');
    setDescription('');
    setImages([]);
    setCodeSnippets([]);
    setTimeLimit(undefined);
    setAddingImage(false);
    setAddingCodeSnippet(false);
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
    { key: QuestionType.ShortAnswer, text: 'Short Answer' },
    { key: QuestionType.Matching, text: 'Matching' }
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
          <div className={customStyles.choicesContainer}>
            <Text styles={formSectionTitleStyles}>Choices (select the correct answer)</Text>
            {choices.map((choice, idx) => (
              <div key={choice.id} className={styles.choiceRow}>
                <Stack horizontal tokens={choiceRowStackTokens} verticalAlign="center" style={{ width: '100%' }}>
                  <Stack.Item>
                    <Checkbox
                      checked={choice.isCorrect}
                      onChange={(_e, checked) => handleChoiceCorrectChange(choice.id, !!checked)}
                      label=""
                      className={styles.choiceCheckbox}
                    />
                  </Stack.Item>
                  <Stack.Item grow>
                    <TextField
                      placeholder={`Choice ${idx + 1}`}
                      value={choice.text}
                      onChange={(_e, value) => handleChoiceTextChange(choice.id, value || '')}
                      className={styles.choiceTextField}
                    />
                  </Stack.Item>
                  <Stack.Item align="center">
                    <IconButton
                      iconProps={{ iconName: 'Picture' }}
                      title="Add Image to Choice"
                      ariaLabel="Add Image to Choice"
                      onClick={() => {
                        // Create an empty image to start with
                        const newImage: IQuizImage = {
                          id: uuidv4(),
                          url: '',
                          fileName: '',
                          altText: `Image for ${choice.text}`
                        };
                        handleChoiceImageUpload(choice.id, newImage);
                      }}
                      styles={{ root: { margin: '0 8px' } }}
                    />
                  </Stack.Item>
                  <Stack.Item>
                    <IconButton
                      iconProps={deleteIcon}
                      title="Delete"
                      ariaLabel="Delete"
                      onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                      disabled={choices.length <= 2} // Require at least 2 choices
                      className={styles.choiceDeleteButton}
                    />
                  </Stack.Item>
                </Stack>

                {choice.image && (
                  <div className={styles.choiceImageContainer}>
                    <ImageUpload
                      currentImage={choice.image}
                      onImageUpload={(image) => handleChoiceImageUpload(choice.id, image)}
                      onImageRemove={() => handleChoiceImageRemove(choice.id)}
                      label={`Image for Choice ${idx + 1}`}
                      context={context}
                    />
                  </div>
                )}
              </div>
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
          <div className={customStyles.choicesContainer}>
            <Text styles={formSectionTitleStyles}>Select the correct answer:</Text>
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
              styles={{ root: { marginTop: 16 } }}
            />
          </div>
        );

      case QuestionType.MultiSelect: {
        return (
          <div className={customStyles.choicesContainer}>
            <Text styles={formSectionTitleStyles}>Choices (select all correct answers)</Text>
            {choices.map((choice, idx) => (
              <div key={choice.id} className={styles.choiceRow}>
                <Stack horizontal tokens={choiceRowStackTokens} verticalAlign="center" style={{ width: '100%' }}>
                  <Stack.Item>
                    <Checkbox
                      checked={choice.isCorrect}
                      onChange={(_e, checked) => handleChoiceCorrectChange(choice.id, !!checked)}
                      label=""
                      className={styles.choiceCheckbox}
                    />
                  </Stack.Item>
                  <Stack.Item grow>
                    <TextField
                      value={choice.text}
                      onChange={(_e, value) => handleChoiceTextChange(choice.id, value || '')}
                      placeholder={`Choice ${idx + 1}`}
                      className={styles.choiceTextField}
                    />
                  </Stack.Item>
                  <Stack.Item align="center">
                    <IconButton
                      iconProps={{ iconName: 'Picture' }}
                      title="Add Image to Choice"
                      ariaLabel="Add Image to Choice"
                      onClick={() => {
                        // Create an empty image to start with
                        const newImage: IQuizImage = {
                          id: uuidv4(),
                          url: '',
                          fileName: '',
                          altText: `Image for ${choice.text}`
                        };
                        handleChoiceImageUpload(choice.id, newImage);
                      }}
                      styles={{ root: { margin: '0 8px' } }}
                    />
                  </Stack.Item>
                  <Stack.Item>
                    <IconButton
                      iconProps={deleteIcon}
                      title="Delete"
                      ariaLabel="Delete"
                      onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                      disabled={choices.length <= 2} // Require at least 2 choices
                      className={styles.choiceDeleteButton}
                    />
                  </Stack.Item>
                </Stack>

                {choice.image && (
                  <div className={styles.choiceImageContainer}>
                    <ImageUpload
                      currentImage={choice.image}
                      onImageUpload={(image) => handleChoiceImageUpload(choice.id, image)}
                      onImageRemove={() => handleChoiceImageRemove(choice.id)}
                      label={`Image for Choice ${idx + 1}`}
                      context={context}
                    />
                  </div>
                )}
              </div>
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
          <div className={customStyles.choicesContainer}>
            <Text styles={formSectionTitleStyles}>Correct Answer</Text>
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
              styles={{ root: { marginTop: 16 } }}
            />
          </div>
        );
      case QuestionType.Matching: {
        return (
          <div className={customStyles.matchingContainer}>
            <Text styles={formSectionTitleStyles}>Create Matching Pairs</Text>

            {matchingPairs.map((pair, idx) => (
              <div key={pair.id} className={styles.matchingPairRow}>
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  <Stack.Item grow>
                    <TextField
                      label={`Left Item ${idx + 1}`}
                      value={pair.leftItem}
                      onChange={(_e, value) => {
                        const updatedPairs = [...matchingPairs];
                        updatedPairs[idx].leftItem = value || '';
                        setMatchingPairs(updatedPairs);
                      }}
                      className={styles.matchingItemField}
                    />
                  </Stack.Item>
                  <Stack.Item grow>
                    <TextField
                      label={`Right Item ${idx + 1}`}
                      value={pair.rightItem}
                      onChange={(_e, value) => {
                        const updatedPairs = [...matchingPairs];
                        updatedPairs[idx].rightItem = value || '';
                        setMatchingPairs(updatedPairs);
                      }}
                      className={styles.matchingItemField}
                    />
                  </Stack.Item>
                  <Stack.Item align="end">
                    <IconButton
                      iconProps={deleteIcon}
                      title="Remove Pair"
                      ariaLabel="Remove Matching Pair"
                      onClick={() => {
                        const updatedPairs = matchingPairs.filter((_, i) => i !== idx);
                        setMatchingPairs(updatedPairs);
                      }}
                      className={styles.matchingDeleteButton}
                      disabled={matchingPairs.length <= 2}
                      styles={{ root: { marginTop: '29px' } }}
                    />
                  </Stack.Item>
                </Stack>
              </div>
            ))}

            <DefaultButton
              iconProps={addIcon}
              text="Add Matching Pair"
              onClick={() => {
                const newPair: IMatchingPair = {
                  id: uuidv4(),
                  leftItem: '',
                  rightItem: ''
                };

                setMatchingPairs([...matchingPairs, newPair]);
              }}
              className={styles.addPairButton}
              styles={{ root: { marginTop: '16px' } }}
            />
          </div>
        );
      }

      default:
        return null;
    }
  };

  // Get the dialog content props
  const dialogContentProps: IDialogContentProps = {
    type: DialogType.largeHeader,
    title: initialQuestion ? 'Edit Question' : 'Add New Question',
    className: customStyles.dialogContent
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={onCancel}
      dialogContentProps={dialogContentProps}
      modalProps={customModalProps}
      className={customStyles.dialogRoot}
    >
      {validationError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          dismissButtonAriaLabel="Close"
          styles={{ root: { marginBottom: 16, marginTop: 8 } }}
        >
          {validationError}
        </MessageBar>
      )}

      <Pivot
        selectedKey={activeTab}
        onLinkClick={(item) => item && setActiveTab(item.props.itemKey || 'question')}
        styles={{
          root: {
            marginBottom: 20,
            position: 'relative',
            zIndex: 10 // Ensure pivot headers appear above rich text editor
          }
        }}
      >
        <PivotItem headerText="Question" itemKey="question" />
        <PivotItem headerText="Additional Info" itemKey="additional" />
      </Pivot>

      {activeTab === 'question' && (
        <div className={customStyles.formContainer}>
          <TextField
            label="Question Title"
            required
            value={title}
            onChange={(_e, value) => setTitle(value || '')}
            placeholder="Enter your question title here"
          />

          <div className={customStyles.sectionContainer}>
            <Text styles={formSectionTitleStyles}>Question Description (Optional)</Text>
            <div className={styles.richTextEditor}>
              <RichText
                value={description}
                onChange={handleDescriptionChange}
                isEditMode={true}
                placeholder="Add detailed question text, instructions, or context here..."
                style={{
                  minHeight: '120px',
                  height: '200px',
                  overflowY: 'visible'
                }}
              />
            </div>
          </div>

          <div className={customStyles.horizontalGroup}>
            <div style={{ flex: '1', minWidth: '250px' }}>
              <Dropdown
                label="Question Type"
                required
                selectedKey={questionType}
                onChange={(_e, option) => {
                  if (option) {
                    const newType = option.key as QuestionType;
                    setQuestionType(newType);

                    // Reset or initialize different fields based on question type
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
                    } else if (newType === QuestionType.Matching) {
                      // Initialize with two empty matching pairs
                      setMatchingPairs([
                        { id: uuidv4(), leftItem: '', rightItem: '' },
                        { id: uuidv4(), leftItem: '', rightItem: '' }
                      ]);
                    }
                  }
                }}
                options={questionTypeOptions}
              />
            </div>

            <div style={{ flex: '1', minWidth: '250px' }}>
              <Dropdown
                label="Category"
                required
                selectedKey={category}
                onChange={(_e, option) => option && setCategory(option.key as string)}
                options={categoryOptions}
                placeholder="Select or add category"
              />
            </div>
          </div>

          {category === 'new' && (
            <TextField
              label="New Category"
              required
              value={newCategory}
              onChange={(_e, value) => setNewCategory(value || '')}
              placeholder="Enter new category"
            />
          )}

          {renderQuestionTypeInputs()}
        </div>
      )}

      {activeTab === 'additional' && (
        <div className={customStyles.formContainer}>
          <div className={customStyles.horizontalGroup}>
            <div style={{ flex: '1', minWidth: '250px' }}>
              <Toggle
                label="Enable Time Limit"
                checked={timeLimitEnabled}
                onChange={(_e, checked) => {
                  setTimeLimitEnabled(!!checked);
                  if (!checked) {
                    setTimeLimit(undefined); // Clear the time limit if disabled
                  } else if (timeLimit === undefined) {
                    // Set a default value when enabling
                    setTimeLimit(60);
                  }
                }}
                onText="On"
                offText="Off"
              />
            </div>

            <div style={{ flex: '0 0 auto', minWidth: '150px' }}>
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
              />
            </div>

            <div style={{ flex: '0 0 auto', minWidth: '200px' }}>
              {timeLimitEnabled ? (
                <SpinButton
                  label="Time Limit (seconds)"
                  labelPosition={0}
                  value={timeLimit !== undefined ? timeLimit.toString() : ''}
                  min={5}
                  max={3600}
                  step={5}
                  onChange={(_e, value) => handleTimeLimitChange(value || '')}
                  incrementButtonAriaLabel="Increase value by 5 seconds"
                  decrementButtonAriaLabel="Decrease value by 5 seconds"
                  styles={spinButtonStyles}
                />
              ) : null}
            </div>
          </div>

          <div className={customStyles.sectionContainer}>
            <Text styles={formSectionTitleStyles}>Explanation (Optional)</Text>
            <TextField
              multiline
              rows={4}
              value={explanation}
              onChange={(_e, value) => setExplanation(value || '')}
              placeholder="Enter an explanation for the correct answer"
            />
          </div>

          {/* Code Snippet Section */}
          <div className={customStyles.sectionContainer}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: '16px' }}>
              <Text styles={formSectionTitleStyles}>Code Snippets</Text>
              {!addingCodeSnippet && (
                <DefaultButton
                  iconProps={codeIcon}
                  text="Add Code Snippet"
                  onClick={() => setAddingCodeSnippet(true)}
                />
              )}
            </Stack>

            {codeSnippets.length > 0 ? (
              <Stack tokens={{ childrenGap: 16 }}>
                {codeSnippets.map(snippet => (
                  <CodeSnippet
                    key={snippet.id}
                    snippet={snippet}
                    onChange={handleCodeSnippetChange}
                    onRemove={() => handleCodeSnippetRemove(snippet.id)}
                  />
                ))}
              </Stack>
            ) : (
              <Text style={{ color: '#666', fontStyle: 'italic' }}>No code snippets added yet.</Text>
            )}

            {addingCodeSnippet && (
              <div style={{ marginTop: '16px' }}>
                <CodeSnippet
                  onChange={handleCodeSnippetChange}
                  onRemove={() => setAddingCodeSnippet(false)}
                  isEditing={true}
                />
              </div>
            )}
          </div>

          {/* Images Section */}
          <div className={customStyles.sectionContainer}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: '16px' }}>
              <Text styles={formSectionTitleStyles}>Images</Text>
              {!addingImage && (
                <DefaultButton
                  iconProps={imageIcon}
                  text="Add Image"
                  onClick={() => setAddingImage(true)}
                />
              )}
            </Stack>

            {images.length > 0 ? (
              <Stack tokens={{ childrenGap: 16 }}>
                {images.map(image => (
                  <div key={image.id} className={styles.imageItem}>
                    <ImageUpload
                      currentImage={image}
                      onImageUpload={handleImageUpload}
                      onImageRemove={() => handleImageRemove(image.id)}
                      context={context}
                    />
                  </div>
                ))}
              </Stack>
            ) : (
              <Text style={{ color: '#666', fontStyle: 'italic' }}>No images added yet.</Text>
            )}

            {addingImage && (
              <div style={{ marginTop: '16px' }}>
                <ImageUpload
                  onImageUpload={handleImageUpload}
                  onImageRemove={() => setAddingImage(false)}
                  label="Upload Image"
                  context={context}
                />
              </div>
            )}
          </div>
        </div>
      )}

      <DialogFooter className={customStyles.footer}>
        <Stack 
          horizontal 
          wrap 
          tokens={{ childrenGap: 12 }} 
          horizontalAlign="end"
          verticalAlign="center"
          styles={{ root: { width: '100%' } }}
        >
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
        </Stack>
      </DialogFooter>
    </Dialog>
  );
}
export default AddQuestionDialog;