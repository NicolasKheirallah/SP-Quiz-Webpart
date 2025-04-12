import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  IAddQuestionFormProps,
  IChoice,
  QuestionType,
  IQuizQuestion,
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
// const timerIcon: IIconProps = { iconName: 'Timer' };

// Stack token for spacing
const stackTokens: IStackTokens = {
  childrenGap: 15
};

// Custom styles for the form section title
const formSectionTitleStyles: ITextStyles = {
  root: {
    fontSize: '14px',
    fontWeight: 600 as const,
    marginBottom: '8px',
    color: '#323130'
  }
};

// Spin button styles
const spinButtonStyles: Partial<ISpinButtonStyles> = {
  spinButtonWrapper: {
    width: 100
  }
};

const AddQuestionDialog: React.FC<IAddQuestionFormProps> = ({
  categories,
  onSubmit,
  onCancel,
  isSubmitting,
  onPreviewQuestion,
  initialQuestion,
  defaultQuestionTimeLimit,
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
    setCorrectChoiceId('');
    setShortAnswerText('');
    setExplanation('');
    setPoints('1');
    setCaseSensitive(false);
    setValidationError('');
    setDescription('');
    // Reset new feature fields
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
            <Text styles={formSectionTitleStyles}>Choices (select the correct answer)</Text>
            {choices.map((choice, idx) => (
              <div key={choice.id} className={styles.choiceRow}>
                <div className={styles.choiceInputRow}>
                  <Checkbox
                    checked={choice.isCorrect}
                    onChange={(_e, checked) => handleChoiceCorrectChange(choice.id, !!checked)}
                    label=""
                    className={styles.choiceCheckbox}
                  />
                  <TextField
                    placeholder={`Choice ${idx + 1}`}
                    value={choice.text}
                    onChange={(_e, value) => handleChoiceTextChange(choice.id, value || '')}
                    className={styles.choiceTextField}
                  />
                  <IconButton
                    iconProps={deleteIcon}
                    title="Delete"
                    ariaLabel="Delete"
                    onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                    disabled={choices.length <= 2} // Require at least 2 choices
                    className={styles.choiceDeleteButton}
                  />
                </div>

                {choice.image ? (
                  <div className={styles.choiceImageContainer}>
                    <ImageUpload
                      currentImage={choice.image}
                      onImageUpload={(image) => handleChoiceImageUpload(choice.id, image)}
                      onImageRemove={() => handleChoiceImageRemove(choice.id)}
                      label={`Image for Choice ${idx + 1}`}
                      context={context}
                    />
                  </div>
                ) : (
                  <DefaultButton
                    iconProps={imageIcon}
                    text="Add Image to Choice"
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
                    className={styles.addImageChoiceButton}
                  />
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
          <div className={styles.choicesContainer}>
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
            />
          </div>
        );

      case QuestionType.MultiSelect: {
        return (
          <div className={styles.choicesContainer}>
            <Text styles={formSectionTitleStyles}>Choices (select all correct answers)</Text>
            {choices.map((choice, idx) => (
              <div key={choice.id} className={styles.choiceRow}>
                <div className={styles.choiceInputRow}>
                  <Checkbox
                    checked={choice.isCorrect}
                    onChange={(_e, checked) => handleChoiceCorrectChange(choice.id, !!checked)}
                    label=""
                    className={styles.choiceCheckbox}
                  />
                  <TextField
                    value={choice.text}
                    onChange={(_e, value) => handleChoiceTextChange(choice.id, value || '')}
                    placeholder={`Choice ${idx + 1}`}
                    className={styles.choiceTextField}
                  />
                  <IconButton
                    iconProps={deleteIcon}
                    title="Delete"
                    ariaLabel="Delete"
                    onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                    disabled={choices.length <= 2} // Require at least 2 choices
                    className={styles.choiceDeleteButton}
                  />
                </div>

                {choice.image ? (
                  <div className={styles.choiceImageContainer}>
                    <ImageUpload
                      currentImage={choice.image}
                      onImageUpload={(image) => handleChoiceImageUpload(choice.id, image)}
                      onImageRemove={() => handleChoiceImageRemove(choice.id)}
                      label={`Image for Choice ${idx + 1}`}
                      context={context}
                    />
                  </div>
                ) : (
                  <DefaultButton
                    iconProps={imageIcon}
                    text="Add Image to Choice"
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
                    className={styles.addImageChoiceButton}
                  />
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
          <div className={styles.choicesContainer}>
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
              styles={{ root: { marginTop: 12 } }}
            />
          </div>
        );

      default:
        return null;
    }
  };

  // Get the dialog content and modal props
  const dialogContentProps: IDialogContentProps = {
    type: DialogType.largeHeader,
    title: initialQuestion ? 'Edit Question' : 'Add New Question'
  };

  const modalProps: IModalProps = {
    isBlocking: true,
    styles: {
      main: {
        minWidth: '320px',
        maxWidth: '850px',
        width: '90vw'
      }
    }
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={onCancel}
      dialogContentProps={dialogContentProps}
      modalProps={modalProps}
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
        <Stack tokens={stackTokens} className={styles.formWrapper}>
          <TextField
            label="Question Title"
            required
            value={title}
            onChange={(_e, value) => setTitle(value || '')}
            placeholder="Enter your question title here"
            styles={{ fieldGroup: { width: '100%' } }}
          />

          <div className={styles.formSection}>
            <Text styles={formSectionTitleStyles}>Question Description (Optional)</Text>
            <div className={styles.richTextEditor}>
              <RichText
                value={description}
                onChange={handleDescriptionChange}
                isEditMode={true}
                placeholder="Add detailed question text, instructions, or context here..."
                style={{
                  minHeight: '120px'
                }}
              />
            </div>
          </div>

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
          <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
            <Stack.Item>
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
            </Stack.Item>

            <Stack.Item>
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
            </Stack.Item>
          </Stack>

          <TextField
            label="Explanation (Optional)"
            multiline
            rows={4}
            value={explanation}
            onChange={(_e, value) => setExplanation(value || '')}
            placeholder="Enter an explanation for the correct answer"
            styles={{ fieldGroup: { width: '100%' } }}
          />

          {/* Code Snippet Section */}
          <Stack tokens={stackTokens}>
            <Text styles={formSectionTitleStyles}>Code Snippets</Text>

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
              <Text>No code snippets added yet.</Text>
            )}

            {addingCodeSnippet ? (
              <CodeSnippet
                onChange={handleCodeSnippetChange}
                onRemove={() => setAddingCodeSnippet(false)}
                isEditing={true}
              />
            ) : (
              <DefaultButton
                iconProps={codeIcon}
                text="Add Code Snippet"
                onClick={() => setAddingCodeSnippet(true)}
                styles={{ root: { marginTop: 8 } }}
              />
            )}
          </Stack>

          {/* Images Section */}
          <Stack tokens={stackTokens}>
            <Text styles={formSectionTitleStyles}>Images</Text>

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
              <Text>No images added yet.</Text>
            )}

            {addingImage ? (
              <ImageUpload
                onImageUpload={handleImageUpload}
                onImageRemove={() => setAddingImage(false)}
                label="Upload Image"
                context={context}
              />
            ) : (
              <DefaultButton
                iconProps={imageIcon}
                text="Add Image"
                onClick={() => setAddingImage(true)}
                styles={{ root: { marginTop: 8 } }}
              />
            )}
          </Stack>
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