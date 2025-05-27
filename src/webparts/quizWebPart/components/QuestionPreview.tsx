import * as React from 'react';
import { 
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  Label,
  Text,
  Stack,
  IStackTokens,
  Checkbox,
  TextField,
  MessageBar,
  MessageBarType,
  ChoiceGroup,
  IChoiceGroupOption,
  Icon
} from '@fluentui/react';
import { IQuestionPreviewProps, QuestionType } from './interfaces';
import styles from './Quiz.module.scss';

// Stack tokens for spacing
const stackTokens: IStackTokens = {
  childrenGap: 10
};

const QuestionPreview: React.FC<IQuestionPreviewProps> = (props) => {
  const { question, onClose } = props;
  
  // Helper function to render different question types
  const renderQuestionContent = (): JSX.Element => {
    switch (question.type) {
      case QuestionType.MultipleChoice:
      case QuestionType.TrueFalse: {
        const options: IChoiceGroupOption[] = question.type === QuestionType.MultipleChoice
          ? question.choices.map(choice => ({
              key: choice.id,
              text: choice.text,
              disabled: true
            }))
          : [
              { key: 'true', text: 'True', disabled: true },
              { key: 'false', text: 'False', disabled: true }
            ];
        
        return (
          <Stack tokens={stackTokens}>
            <ChoiceGroup
              options={options}
              selectedKey={question.choices.find(c => c.isCorrect)?.id}
            />
            {question.choices.filter(c => c.isCorrect).map(correctChoice => (
              <div 
                key={correctChoice.id} 
                className={`${styles.statusBar} ${styles.success}`}
              >
                <Stack horizontal tokens={stackTokens} verticalAlign="center">
                  <Icon iconName="CheckMark" style={{ color: 'green', marginRight: '8px' }} />
                  <Text>Correct Answer: {correctChoice.text}</Text>
                </Stack>
              </div>
            ))}
          </Stack>
        );
      }
        
      case QuestionType.MultiSelect:
        return (
          <Stack tokens={stackTokens}>
            {question.choices.map((choice) => (
              <Stack key={choice.id} className={choice.isCorrect ? `${styles.statusBar} ${styles.success}` : styles.choiceRow}>
                <Stack horizontal tokens={stackTokens} verticalAlign="center">
                  <Checkbox
                    disabled
                    checked={choice.isCorrect}
                    label={choice.text}
                  />
                  {choice.isCorrect && <Icon iconName="CheckMark" style={{ color: 'green', marginLeft: '8px' }} />}
                </Stack>
              </Stack>
            ))}
          </Stack>
        );
        
      case QuestionType.ShortAnswer:
        return (
          <Stack tokens={stackTokens}>
            <TextField
              readOnly
              placeholder="[Student answer will appear here]"
              borderless
              styles={{ fieldGroup: { background: 'white', padding: '8px' } }}
            />
            <div className={`${styles.statusBar} ${styles.success}`}>
              <Label>Correct Answer:</Label>
              <Text>{question.correctAnswer}</Text>
              {question.caseSensitive && (
                <MessageBar
                  messageBarType={MessageBarType.info}
                  styles={{ root: { marginTop: '8px' } }}
                >
                  Answers are case sensitive
                </MessageBar>
              )}
            </div>
          </Stack>
        );

      case QuestionType.Matching:
        return (
          <Stack tokens={stackTokens}>
            <Label>Matching Pairs:</Label>
            {question.matchingPairs && question.matchingPairs.length > 0 ? (
              <Stack tokens={stackTokens}>
                {question.matchingPairs.map((pair, index) => (
                  <div key={pair.id} className={`${styles.statusBar} ${styles.success}`}>
                    <Stack horizontal tokens={stackTokens} verticalAlign="center">
                      <Icon iconName="CheckMark" style={{ color: 'green', marginRight: '8px' }} />
                      <Text>{pair.leftItem} → {pair.rightItem}</Text>
                    </Stack>
                  </div>
                ))}
              </Stack>
            ) : (
              <MessageBar messageBarType={MessageBarType.warning}>
                No matching pairs configured for this question.
              </MessageBar>
            )}
            <MessageBar
              messageBarType={MessageBarType.info}
              styles={{ root: { marginTop: '8px' } }}
            >
              Students will need to match each left item with its corresponding right item.
            </MessageBar>
          </Stack>
        );
        
      default:
        return (
          <Text>Preview not available for this question type.</Text>
        );
    }
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={onClose}
      dialogContentProps={{
        type: DialogType.normal,
        title: 'Question Preview'
      }}
      modalProps={{
        isBlocking: false,
        styles: { main: { maxWidth: '600px' } }
      }}
    >
      <div className={styles.formWrapper}>
        <div className={styles.questionCard}>
          <Text className={styles.questionTitle}>{question.title}</Text>
          
          <div className={styles.questionInfo}>
            <span className={styles.questionTag}>Category: {question.category}</span>
            <span className={styles.questionTag}>Type: {question.type}</span>
            {question.points && (
              <span className={styles.questionTag}>{question.points} points</span>
            )}
          </div>
          
          <Stack className={styles.formSection}>
            {renderQuestionContent()}
          </Stack>
          
          {question.explanation && (
            <div className={`${styles.statusBar} ${styles.warning}`}>
              <Stack horizontal tokens={{ childrenGap: 5 }} verticalAlign="center">
                <Icon iconName="Info" style={{ color: '#8f6f00' }} />
                <Label styles={{ root: { marginBottom: 0, color: '#8f6f00' } }}>Explanation:</Label>
              </Stack>
              <Text>{question.explanation}</Text>
            </div>
          )}
        </div>
      </div>
      
      <DialogFooter className={styles.formButtons}>
        <PrimaryButton onClick={onClose} text="Close" />
      </DialogFooter>
    </Dialog>
  );
};

export default QuestionPreview;