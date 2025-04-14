import * as React from 'react';
import { IQuizResultsProps } from './interfaces';
import {
  Stack,
  IStackTokens,
  Text,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  ITextStyles,
  FontWeights,
  IIconProps,
  Pivot,
  PivotItem,
  Separator,
  Label
} from '@fluentui/react';
import { Accordion } from "@pnp/spfx-controls-react/lib/Accordion";
import styles from './Quiz.module.scss';

const retakeIcon: IIconProps = { iconName: 'Refresh' };
const reviewIcon: IIconProps = { iconName: 'ReadingMode' };
const excellentIconName = 'Trophy';
const goodIconName = 'Emoji2';
const averageIconName = 'EmojiNeutral';
const poorIconName = 'Sad';
const checkIconName = 'CheckMark';
const questionIcon: IIconProps = { iconName: 'QuestionAnswer' };

// Styles
const resultsTitleStyles: ITextStyles = {
  root: {
    fontSize: '28px',
    fontWeight: FontWeights.semibold,
    marginBottom: '24px',
    textAlign: 'center',
    color: '#0078d4'
  }
};

const scoreValueStyles: ITextStyles = {
  root: {
    fontSize: '42px',
    fontWeight: FontWeights.bold,
    color: 'white',
    margin: 0
  }
};

const scoreTextStyles: ITextStyles = {
  root: {
    fontSize: '16px',
    color: 'white',
    marginTop: '4px',
    opacity: 0.9
  }
};

const scoreDetailsStyles: ITextStyles = {
  root: {
    fontSize: '18px',
    marginBottom: '16px'
  }
};

const resultMessageStyles: ITextStyles = {
  root: {
    fontSize: '18px',
    fontStyle: 'italic',
    marginBottom: '16px',
    textAlign: 'center'
  }
};

// Get appropriate icon based on score
const getScoreIconName = (percentage: number): string => {
  if (percentage >= 90) return excellentIconName;
  if (percentage >= 70) return goodIconName;
  if (percentage >= 50) return averageIconName;
  return poorIconName;
};

const QuizResults: React.FC<IQuizResultsProps> = (props) => {
  const {
    score,
    totalQuestions,
    totalPoints,
    isSubmitting,
    submissionSuccess,
    submissionError,
    onRetakeQuiz,
    messages,
    detailedResults
  } = props;

  const [activeView, setActiveView] = React.useState<string>('summary');

  // Calculate the percentage of correctly answered questions vs total questions
  const correctQuestionsCount = detailedResults?.correctlyAnsweredQuestions || 0;
  const correctVsTotalPercentage = totalQuestions > 0 
    ? Math.round((correctQuestionsCount / totalQuestions) * 100) 
    : 0;

  // Determine result message based on correctVsTotalPercentage
  let resultMessage = '';
  if (correctVsTotalPercentage >= 90) {
    resultMessage = messages.excellent || 'Excellent! You have mastered this topic!';
  } else if (correctVsTotalPercentage >= 70) {
    resultMessage = messages.good || 'Good job! You have a solid understanding.';
  } else if (correctVsTotalPercentage >= 50) {
    resultMessage = messages.average || 'Not bad. There\'s room for improvement.';
  } else {
    resultMessage = messages.poor || 'Keep studying. You\'ll get better with practice.';
  }

  // Stack tokens for spacing
  const stackTokens: IStackTokens = {
    childrenGap: 16
  };

  // Format answers for display - ensure we're displaying text, not IDs
  const formatAnswer = (answer: string | string[] | undefined): string => {
    if (answer === undefined || answer === null) return 'No answer provided';

    if (Array.isArray(answer)) {
      return answer.length > 0 ? answer.join(', ') : 'No answer provided';
    }
    
    // Check if answer looks like a UUID (simple check)
    const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (typeof answer === 'string' && uuidPattern.test(answer)) {
      return 'ID format detected - please check data conversion';
    }
    
    return answer.toString().trim() !== '' ? answer.toString() : 'No answer provided';
  };

  // Render summary view
  const renderSummaryView = (): JSX.Element => {
    return (
      <div className={styles.scoreCard}>
        <div className={styles.resultsSummaryCard}>
          <Stack horizontal wrap horizontalAlign="space-between" tokens={{ childrenGap: 20 }}>
            <Stack style={{ flex: '1 1 auto', minWidth: '250px' }}>
              <div className={styles.resultsStat}>
                <Icon iconName={questionIcon.iconName} style={{ color: '#0078d4', marginRight: '12px', fontSize: '20px' }} />
                <Text styles={scoreDetailsStyles}>
                  You answered <b>{correctQuestionsCount}</b> out of <b>{totalQuestions}</b> questions correctly.
                </Text>
              </div>
              <div className={styles.resultsStat}>
                <Text styles={scoreDetailsStyles}>
                  Points earned: <b>{score}</b> out of <b>{totalPoints}</b> points.
                </Text>
              </div>
            </Stack>
            
            <div className={styles.scoreBadge} style={{
              backgroundColor: correctVsTotalPercentage >= 90 ? '#107c10' :
                correctVsTotalPercentage >= 70 ? '#498205' :
                  correctVsTotalPercentage >= 50 ? '#ffaa44' : '#d13438',
              width: '120px',
              height: '120px'
            }}>
              <Text styles={scoreValueStyles}>{correctVsTotalPercentage}%</Text>
              <Text styles={scoreTextStyles}>Score</Text>
            </div>
          </Stack>
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center" verticalAlign="center" style={{ margin: '16px 0' }}>
          <Icon iconName={getScoreIconName(correctVsTotalPercentage)} styles={{ root: { fontSize: '24px' } }} />
          <Text styles={resultMessageStyles}>{resultMessage}</Text>
        </Stack>

        <Stack horizontalAlign="center" tokens={stackTokens} style={{ marginTop: '10px' }}>
          <ProgressIndicator
            percentComplete={correctVsTotalPercentage / 100}
            styles={{
              root: { maxWidth: '400px', margin: '20px auto 0' },
              itemName: { textAlign: 'left' },
              itemProgress: { paddingBottom: 4 }
            }}
          />
        </Stack>
      </div>
    );
  };

  // Render detailed results view
  const renderDetailedView = (): JSX.Element => {
    if (!detailedResults || !detailedResults.questionResults) {
      return (
        <MessageBar messageBarType={MessageBarType.info}>
          Detailed results are not available or loading.
        </MessageBar>
      );
    }

    return (
      <div className={styles.resultsContainer}>
        {/* Summary Stats Card */}
        <div className={styles.resultsSummaryCard}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack tokens={{ childrenGap: 5 }}>
              <Text variant="large" style={{ fontWeight: FontWeights.semibold, color: '#0078d4' }}>
                {correctQuestionsCount} / {totalQuestions} questions correct ({correctVsTotalPercentage}%)
              </Text>
              <Text variant="medium">
                Points: {score} / {totalPoints}
              </Text>
            </Stack>
            <div className={styles.scoreCircle} style={{
              backgroundColor: correctVsTotalPercentage >= 90 ? '#107c10' :
                correctVsTotalPercentage >= 70 ? '#498205' :
                  correctVsTotalPercentage >= 50 ? '#ffaa44' : '#d13438'
            }}>
              <div className={styles.scoreValue}>{correctVsTotalPercentage}%</div>
              <div className={styles.scoreLabel}>Score</div>
            </div>
          </Stack>
        </div>

        {/* Enhanced Results Table */}
        <table className={styles.resultsTable}>
          <thead>
            <tr>
              <th>Question</th>
              <th>Result</th>
              <th>Points</th>
            </tr>
          </thead>
          <tbody>
            {detailedResults.questionResults.map((question, index) => (
              <tr key={question.id || index}>
                <td>{question.title || `Question ${index + 1}`}</td>
                <td>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                    <Icon
                      iconName={question.isCorrect ? 'CheckMark' : 'Cancel'}
                      style={{
                        color: question.isCorrect ? '#107c10' : '#d13438',
                        fontSize: '16px'
                      }}
                    />
                    <Text>{question.isCorrect ? 'Correct' : 'Incorrect'}</Text>
                  </Stack>
                </td>
                <td>{question.earnedPoints} / {question.points}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <Separator styles={{
          root: {
            height: 2,
            backgroundColor: '#0078d4',
            marginTop: '24px'
          }
        }}>Question Details</Separator>

        {/* Improved Accordion Design */}
        <div style={{ marginTop: '16px' }}>
          {detailedResults.questionResults.map((question, index) => {
            if (!question) return null;

            const simpleAccordionTitle = `Question ${index + 1}: ${question.title || 'Untitled Question'}`;

            const accordionContent = (
              <div className={styles.questionDetails}>
                <div className={`${styles.resultStatusBox} ${question.isCorrect ? styles.correct : styles.incorrect}`}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Icon
                        iconName={question.isCorrect ? 'CheckMark' : 'Cancel'}
                        className={question.isCorrect ? styles.correctIcon : styles.incorrectIcon}
                      />
                      <Text variant="mediumPlus" style={{ fontWeight: FontWeights.semibold }}>
                        Result: {question.isCorrect ? 'Correct' : 'Incorrect'}
                      </Text>
                    </Stack>
                    <Text>Points: {question.earnedPoints} / {question.points}</Text>
                  </Stack>
                </div>

                <Stack horizontal tokens={{ childrenGap: 20 }} wrap style={{ marginTop: '16px' }}>
                  <Stack className={styles.detailColumn}>
                    <Label style={{ color: '#0078d4' }}>Your Answer:</Label>
                    <div className={styles.answerBox}>
                      <Text>{formatAnswer(question.userAnswer)}</Text>
                    </div>
                  </Stack>

                  <Stack className={styles.detailColumn}>
                    <Label style={{ color: '#0078d4' }}>Correct Answer:</Label>
                    <div className={styles.correctAnswerBox}>
                      <Text>{formatAnswer(question.correctAnswer)}</Text>
                    </div>
                  </Stack>
                </Stack>

                {question.explanation && (
                  <div className={styles.explanationBox}>
                    <Label style={{ color: '#0078d4', marginBottom: '4px' }}>Explanation:</Label>
                    <Text>{question.explanation}</Text>
                  </div>
                )}
              </div>
            );

            return (
              <div className={styles.questionAccordion} key={question.id || index}>
                <Accordion
                  title={simpleAccordionTitle}
                  defaultCollapsed={true}
                  className={question.isCorrect ? styles.correctQuestion : styles.incorrectQuestion}
                >
                  {accordionContent}
                </Accordion>
              </div>
            );
          })}
        </div>
      </div>
    );
  };

  // Main return statement
  return (
    <div className={styles.resultsContainer}>
      <Text styles={resultsTitleStyles}>Quiz Results</Text>

      {isSubmitting ? (
        <Stack horizontalAlign="center" tokens={stackTokens} style={{ padding: '30px' }}>
          <Spinner size={SpinnerSize.large} label="Submitting your results..." />
        </Stack>
      ) : (
        <>
          {submissionSuccess && (
            <MessageBar
              messageBarType={MessageBarType.success}
              isMultiline={false}
              styles={{ root: { marginBottom: '20px' } }}
            >
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                <Icon iconName={checkIconName} />
                <span>{messages.success || 'Your score has been successfully recorded!'}</span>
              </Stack>
            </MessageBar>
          )}

          {submissionError && (
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
              styles={{ root: { marginBottom: '20px' } }}
            >
              {submissionError}
            </MessageBar>
          )}

          {detailedResults ? (
            <Pivot
              selectedKey={activeView}
              onLinkClick={(item) => item && setActiveView(item.props.itemKey || 'summary')}
              styles={{ root: { marginBottom: '20px' } }}
              headersOnly={false}
            >
              <PivotItem headerText="Summary" itemKey="summary">
                {renderSummaryView()}
              </PivotItem>
              <PivotItem headerText="Detailed Results" itemKey="details">
                {renderDetailedView()}
              </PivotItem>
            </Pivot>
          ) : (
            renderSummaryView()
          )}

          <Stack horizontal horizontalAlign="center" tokens={stackTokens} style={{ marginTop: '20px' }}>
            <PrimaryButton
              text="Retake Quiz"
              onClick={onRetakeQuiz}
              iconProps={retakeIcon}
              styles={{ root: { minWidth: '140px' } }}
            />
            {activeView === 'summary' && detailedResults && (
              <DefaultButton
                text="View Detailed Results"
                onClick={() => setActiveView('details')}
                iconProps={reviewIcon}
                styles={{ root: { marginLeft: '8px' } }}
              />
            )}
          </Stack>
        </>
      )}
    </div>
  );
};

export default QuizResults;