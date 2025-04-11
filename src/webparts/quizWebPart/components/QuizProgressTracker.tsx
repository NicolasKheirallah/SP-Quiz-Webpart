import * as React from 'react';
import {
  ProgressIndicator,
  Text,
  Stack,
  IStackTokens,
  mergeStyles,
  FontIcon,
  IProgressIndicatorStyles
} from '@fluentui/react';
import { IQuizProgressTrackerProps } from './interfaces';
import styles from './Quiz.module.scss';

const progressContainerClass = mergeStyles({
  padding: '10px 0 15px 0',
  marginBottom: '15px',
  borderBottom: '1px solid #edebe9'
});

const stackTokens: IStackTokens = {
  childrenGap: 8
};

const progressIndicatorStyles: Partial<IProgressIndicatorStyles> = {
  itemProgress: {
    padding: '0 0 8px 0'
  }
};

const QuizProgressTracker: React.FC<IQuizProgressTrackerProps> = (props) => {
  const { progress, showPercentage = true, showNumbers = true, showIcon = true, showTimer = false } = props;
  const { currentQuestion, totalQuestions, answeredQuestions, percentage, remainingTime } = progress;
  
  // Format the remaining time if available
  const formatTime = (seconds: number): string => {
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = seconds % 60;
    return `${minutes}:${remainingSeconds < 10 ? '0' : ''}${remainingSeconds}`;
  };

  return (
    <div className={progressContainerClass}>
      <Stack horizontal tokens={stackTokens} verticalAlign="center" horizontalAlign="space-between">
        <Stack horizontal tokens={stackTokens} verticalAlign="center">
          {showIcon && (
            <FontIcon 
              iconName="ProgressLoopOuter" 
              className={styles.progressIcon} 
            />
          )}
          <Stack>
            {showNumbers && (
              <Text variant="medium">
                Question {currentQuestion} of {totalQuestions} 
                {answeredQuestions < totalQuestions && ` (${answeredQuestions} answered)`}
              </Text>
            )}
            <ProgressIndicator 
              percentComplete={percentage / 100} 
              barHeight={4}
              styles={progressIndicatorStyles}
            />
          </Stack>
        </Stack>
        
        <Stack horizontal tokens={stackTokens} verticalAlign="center">
          {showPercentage && (
            <Text variant="medium" className={styles.progressPercentage}>
              {percentage}% Complete
            </Text>
          )}
          
          {showTimer && remainingTime !== undefined && (
            <div className={styles.timerContainer}>
              <FontIcon 
                iconName="Timer" 
                className={styles.timerIcon} 
              />
              <Text variant="medium" className={styles.timerText}>
                {formatTime(remainingTime)}
              </Text>
            </div>
          )}
        </Stack>
      </Stack>
    </div>
  );
};

export default QuizProgressTracker;