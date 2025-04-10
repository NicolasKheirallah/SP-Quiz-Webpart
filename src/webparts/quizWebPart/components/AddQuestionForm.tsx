import React, { useState } from 'react';
import { IAddQuestionFormProps, IChoice, QuestionType, IQuizQuestion } from './interfaces';
import { v4 as uuidv4 } from 'uuid';
import {
  Input,
  Dropdown,
  Option,
  Button,
  Label,
  Radio,
  RadioGroup,
  Field,
  MessageBar,
  MessageBarBody,
  Textarea,
  Checkbox,
  Card,
  Text,
  TabList,
  Tab
} from '@fluentui/react-components';
import { 
  DeleteRegular, 
  AddRegular, 
  EyeRegular,
  DocumentRegular,
  QuestionCircleRegular,
  TextNumberFormatRegular,
  ArrowResetRegular
} from '@fluentui/react-icons';
import styles from './Quiz.module.scss';

const AddQuestionForm: React.FC<IAddQuestionFormProps> = ({
  categories,
  onSubmit,
  onCancel,
  isSubmitting,
  onPreviewQuestion
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

  const handleChoiceTextChange = (id: string, text: string): void => {
    setChoices(choices.map(c => (c.id === id ? { ...c, text } : c)));
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
      id: Date.now(),
      title,
      category: category === 'new' ? newCategory : category,
      type: questionType,
      choices: choices.filter(c => c.text.trim()), // Filter out empty choices
      correctAnswer: questionType === QuestionType.ShortAnswer ? shortAnswerText : undefined,
      explanation: explanation.trim() || undefined,
      points: pointsValue
    };

    onSubmit(newQuestion);
  };

  const handlePreview = (): void => {
    // Create question object for preview
    const previewQuestion: IQuizQuestion = {
      id: Date.now(),
      title,
      category: category === 'new' ? newCategory : category,
      type: questionType,
      choices: choices.filter(c => c.text.trim()), // Filter out empty choices
      correctAnswer: questionType === QuestionType.ShortAnswer ? shortAnswerText : undefined,
      explanation: explanation.trim() || undefined,
      points: parseInt(points, 10) || 1
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
    setValidationError('');
  };

  // Render different inputs based on question type
  const renderQuestionTypeInputs = (): JSX.Element | null => {
    switch (questionType) {
      case QuestionType.MultipleChoice:
        return (
          <>
            <Label required>Choices</Label>
            <RadioGroup
              value={correctChoiceId}
              onChange={(_, data) => setCorrectChoiceId(data.value)}
            >
              {choices.map((choice, idx) => (
                <div className={styles.choiceRow} key={choice.id}>
                  <Radio 
                    value={choice.id} 
                    checked={choice.isCorrect}
                    onChange={() => handleChoiceCorrectChange(choice.id, true)}
                  />
                  <Input
                    placeholder={`Choice ${idx + 1}`}
                    value={choice.text}
                    onChange={(_, { value }) => handleChoiceTextChange(choice.id, value)}
                  />
                  <Button
                    icon={<DeleteRegular />}
                    appearance="subtle"
                    onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                    disabled={choices.length <= 2} // Require at least 2 choices
                  />
                </div>
              ))}
            </RadioGroup>

            <Button
              icon={<AddRegular />}
              appearance="secondary"
              onClick={() =>
                setChoices([...choices, { id: uuidv4(), text: '', isCorrect: false }])
              }
            >
              Add Choice
            </Button>
          </>
        );
        
      case QuestionType.TrueFalse:
        return (
          <>
            <Label required>Select the correct answer:</Label>
            <RadioGroup
              value={correctChoiceId}
              onChange={(_, data) => {
                setCorrectChoiceId(data.value);
                // Update the choices array to mark the correct answer
                setChoices([
                  { id: 'true', text: 'True', isCorrect: data.value === 'true' },
                  { id: 'false', text: 'False', isCorrect: data.value === 'false' }
                ]);
              }}
            >
              <Radio value="true" label="True" />
              <Radio value="false" label="False" />
            </RadioGroup>
          </>
        );
        
      case QuestionType.MultiSelect:
        return (
          <>
            <Label required>Choices (select all correct answers)</Label>
            {choices.map((choice, idx) => (
              <div className={styles.choiceRow} key={choice.id}>
                <Checkbox
                  checked={choice.isCorrect}
                  onChange={(e) => handleChoiceCorrectChange(choice.id, e.target.checked)}
                  label={`Choice ${idx + 1}`}
                />
                <Input
                  value={choice.text}
                  onChange={(_, { value }) => handleChoiceTextChange(choice.id, value)}
                />
                <Button
                  icon={<DeleteRegular />}
                  appearance="subtle"
                  onClick={() => setChoices(choices.filter(c => c.id !== choice.id))}
                  disabled={choices.length <= 2} // Require at least 2 choices
                />
              </div>
            ))}

            <Button
              icon={<AddRegular />}
              appearance="secondary"
              onClick={() =>
                setChoices([...choices, { id: uuidv4(), text: '', isCorrect: false }])
              }
            >
              Add Choice
            </Button>
          </>
        );
        
      case QuestionType.ShortAnswer:
        return (
          <Field label="Correct Answer" required>
            <Input
              value={shortAnswerText}
              onChange={(_, { value }) => setShortAnswerText(value)}
              placeholder="Enter the correct answer"
            />
          </Field>
        );
        
      default:
        return null;
    }
  };

  return (
    <div className={styles.addQuestionForm}>
      <Card>
        <Text size={600} weight="semibold">Add New Question</Text>

        <TabList selectedValue={activeTab} onTabSelect={(_, data) => setActiveTab(data.value as string)}>
          <Tab value="question" icon={<QuestionCircleRegular />}>Question</Tab>
          <Tab value="additional" icon={<TextNumberFormatRegular />}>Additional Info</Tab>
        </TabList>

        {validationError && (
          <MessageBar intent="error">
            <MessageBarBody>{validationError}</MessageBarBody>
          </MessageBar>
        )}

        {activeTab === 'question' && (
          <>
            <Field label="Question" required>
              <Input
                value={title}
                onChange={(_, { value }) => setTitle(value)}
                placeholder="Enter your question here"
              />
            </Field>

            <Field label="Question Type" required>
              <Dropdown
                value={questionType}
                onOptionSelect={(_, { optionValue }) => {
                  const newType = optionValue as QuestionType;
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
                }}
              >
                <Option value={QuestionType.MultipleChoice}>Multiple Choice</Option>
                <Option value={QuestionType.TrueFalse}>True/False</Option>
                <Option value={QuestionType.MultiSelect}>Multiple Select</Option>
                <Option value={QuestionType.ShortAnswer}>Short Answer</Option>
              </Dropdown>
            </Field>

            <Field label="Category" required>
              <Dropdown
                value={category}
                onOptionSelect={(_, { optionValue }) => setCategory(optionValue || '')}
                placeholder="Select or add category"
              >
                {categories.map(cat => (
                  <Option key={cat} value={cat}>
                    {cat}
                  </Option>
                ))}
                <Option key="new" value="new">
                  Add new category
                </Option>
              </Dropdown>
            </Field>

            {category === 'new' && (
              <Field label="New Category" required>
                <Input
                  value={newCategory}
                  onChange={(_, { value }) => setNewCategory(value)}
                  placeholder="Enter new category"
                />
              </Field>
            )}

            {renderQuestionTypeInputs()}
          </>
        )}

        {activeTab === 'additional' && (
          <>
            <Field label="Points">
              <Input
                value={points}
                onChange={(_, { value }) => {
                  // Ensure only positive numbers are entered
                  const numericValue = value.replace(/[^0-9]/g, '');
                  setPoints(numericValue || '1');
                }}
                type="number"
                min="1"
                placeholder="1"
              />
            </Field>

            <Field label="Explanation (Optional)">
              <Textarea
                value={explanation}
                onChange={(_, { value }) => setExplanation(value)}
                placeholder="Enter an explanation for the correct answer"
                resize="vertical"
              />
            </Field>
          </>
        )}

        <div className={styles.formButtons}>
          <Button
            appearance="primary"
            onClick={handleSubmit}
            disabled={isSubmitting}
            icon={<DocumentRegular />}
          >
            Save Question
          </Button>
          <Button
            appearance="secondary"
            onClick={handlePreview}
            icon={<EyeRegular />}
          >
            Preview
          </Button>
          <Button
            appearance="subtle"
            onClick={resetForm}
            icon={<ArrowResetRegular />}
          >
            Reset
          </Button>
          <Button
            appearance="secondary"
            onClick={onCancel}
            disabled={isSubmitting}
          >
            Cancel
          </Button>
        </div>
      </Card>
    </div>
  );
};

export default AddQuestionForm;