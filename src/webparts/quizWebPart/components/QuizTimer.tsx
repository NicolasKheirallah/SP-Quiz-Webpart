import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import {
  ProgressIndicator,
  Stack,
  Text,
  FontIcon,
  MessageBar,
  MessageBarType,
  IStackTokens
} from '@fluentui/react';
import styles from './Quiz.module.scss';

export interface IQuizTimerProps {
  timeLimit: number; // in seconds
  onTimeExpired: () => void;
  paused: boolean;
  warningThreshold?: number; // percentage of time remaining (default: 20%)
  criticalThreshold?: number; // percentage of time remaining (default: 10%)
}

// Color constants
const NORMAL_COLOR = '#0078d4'; // Blue
const WARNING_COLOR = '#ffaa44'; // Orange
const CRITICAL_COLOR = '#d13438'; // Red

// Stack tokens
const stackTokens: IStackTokens = {
  childrenGap: 8
};

const QuizTimer: React.FC<IQuizTimerProps> = ({
  timeLimit,
  onTimeExpired,
  paused,
  warningThreshold = 20,
  criticalThreshold = 10
}) => {
  const [timeRemaining, setTimeRemaining] = useState<number>(timeLimit);
  const [isWarning, setIsWarning] = useState<boolean>(false);
  const [isCritical, setIsCritical] = useState<boolean>(false);
  const [showNotification, setShowNotification] = useState<boolean>(false);
  
  // Format time as MM:SS or HH:MM:SS if over an hour
  const formatTime = useCallback((seconds: number): string => {
    if (seconds < 0) seconds = 0;
    
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const remainingSeconds = Math.floor(seconds % 60);
    
    if (hours > 0) {
      return `${hours}:${minutes < 10 ? '0' : ''}${minutes}:${remainingSeconds < 10 ? '0' : ''}${remainingSeconds}`;
    }
    
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
  
  // Timer effect
  useEffect(() => {
    if (paused) return;
    
    const timer = setInterval(() => {
      setTimeRemaining(prevTime => {
        const newTime = Math.max(0, prevTime - 1);
        
        // Calculate percentage remaining
        const percentage = calculatePercentageRemaining(newTime);
        
        // Set warning/critical states
        const newIsWarning = percentage <= warningThreshold;
        const newIsCritical = percentage <= criticalThreshold;
        
        // Update states
        setIsWarning(newIsWarning);
        setIsCritical(newIsCritical);
        
        // Show notification when warning or critical threshold is first crossed
        if (newIsWarning && !isWarning) {
          setShowNotification(true);
          setTimeout(() => setShowNotification(false), 5000);
        }
        
        // Handle expiration
        if (newTime === 0) {
          clearInterval(timer);
          onTimeExpired();
        }
        
        return newTime;
      });
    }, 1000);
    
    return () => clearInterval(timer);
  }, [paused, timeLimit, isWarning, isCritical, onTimeExpired, calculatePercentageRemaining, warningThreshold, criticalThreshold]);
  
  // Reset timer if timeLimit changes
  useEffect(() => {
    setTimeRemaining(timeLimit);
  }, [timeLimit]);
  
  const percentageRemaining = calculatePercentageRemaining(timeRemaining);
  const timerColor = getTimerColor(percentageRemaining);
  
  return (
    <div className={styles.quizTimerContainer}>
      {showNotification && (
        <MessageBar
          messageBarType={isCritical ? MessageBarType.severeWarning : MessageBarType.warning}
          isMultiline={false}
          onDismiss={() => setShowNotification(false)}
          dismissButtonAriaLabel="Close"
          className={styles.timerNotification}
        >
          {isCritical 
            ? `Time is running out! Less than ${criticalThreshold}% of time remaining.`
            : `${Math.round(percentageRemaining)}% of time remaining.`
          }
        </MessageBar>
      )}
      
      <Stack horizontal tokens={stackTokens} verticalAlign="center">
        <FontIcon 
          iconName="Timer" 
          className={styles.quizTimerIcon}
          style={{ color: timerColor }}
        />
        
        <Text 
          variant="large"
          className={styles.quizTimerText}
          style={{ 
            color: timerColor, 
            fontWeight: isCritical ? 600 : 400,
            animation: isCritical ? 'pulse 1s infinite' : 'none'
          }}
        >
          {formatTime(timeRemaining)}
        </Text>
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
    </div>
  );
};

export default QuizTimer;