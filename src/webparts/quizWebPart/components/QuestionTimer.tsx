import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { 
  ProgressIndicator, 
  Stack, 
  Text, 
  IStackTokens, 
  FontIcon
} from '@fluentui/react';
import { IQuestionTimerProps } from './interfaces';
import styles from './Quiz.module.scss';

// Stack tokens
const stackTokens: IStackTokens = {
  childrenGap: 8
};

// Color constants
const NORMAL_COLOR = '#0078d4'; // Blue
const WARNING_COLOR = '#ffaa44'; // Orange
const CRITICAL_COLOR = '#d13438'; // Red

const QuestionTimer: React.FC<IQuestionTimerProps> = (props) => {
  const { 
    timeLimit,
    onTimeExpired,
    paused = false, 
    warningThreshold = 20,
    criticalThreshold = 10,
    showText = true
  } = props;
  
  const [timeRemaining, setTimeRemaining] = useState<number>(timeLimit);
  const [isCritical, setIsCritical] = useState<boolean>(false);
  
  // Format time as MM:SS
  const formatTime = useCallback((seconds: number): string => {
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = Math.floor(seconds % 60);
    return `${minutes}:${remainingSeconds < 10 ? '0' : ''}${remainingSeconds}`;
  }, []);
  
  // Calculate percentage remaining
  const calculatePercentageRemaining = useCallback((seconds: number): number => {
    return Math.max(0, Math.min(100, (seconds / timeLimit) * 100));
  }, [timeLimit]);
  
  // Determine color based on time remaining
  const getTimerColor = useCallback((percentage: number): string => {
    if (percentage <= criticalThreshold) return CRITICAL_COLOR;
    if (percentage <= warningThreshold) return WARNING_COLOR;
    return NORMAL_COLOR;
  }, [warningThreshold, criticalThreshold]);
  
  // Update timer state
  useEffect(() => {
    if (paused) return;
    
    const timer = setInterval(() => {
      setTimeRemaining(prevTime => {
        const newTime = Math.max(0, prevTime - 1);
        
        // Check thresholds
        const percentage = calculatePercentageRemaining(newTime);
        setIsCritical(percentage <= criticalThreshold);
        
        // Handle time expiration
        if (newTime === 0) {
          clearInterval(timer);
          onTimeExpired();
        }
        
        return newTime;
      });
    }, 1000);
    
    return () => clearInterval(timer);
  }, [paused, onTimeExpired, calculatePercentageRemaining, criticalThreshold]);
  
  // Reset timer if timeLimit changes
  useEffect(() => {
    setTimeRemaining(timeLimit);
  }, [timeLimit]);
  
  const percentageRemaining = calculatePercentageRemaining(timeRemaining);
  const timerColor = getTimerColor(percentageRemaining);
  
  return (
    <div className={styles.questionTimer}>
      <Stack horizontal tokens={stackTokens} verticalAlign="center">
        <FontIcon 
          iconName="Timer" 
          className={styles.timerIcon}
          style={{ color: timerColor }}
        />
        
        {showText && (
          <Text 
            className={styles.timerText}
            style={{ color: timerColor, fontWeight: isCritical ? 600 : 400 }}
          >
            {formatTime(timeRemaining)}
          </Text>
        )}
      </Stack>
      
      <ProgressIndicator 
        percentComplete={percentageRemaining / 100}
        barHeight={4}
        styles={{ 
          itemProgress: { padding: 0 },
          progressBar: { 
            backgroundColor: timerColor
          }
        }}
      />
      
      {isCritical && (
        <Text 
          variant="small" 
          className={styles.timerWarning}
          style={{ color: CRITICAL_COLOR }}
        >
          Time is running out!
        </Text>
      )}
    </div>
  );
};

export default QuestionTimer;