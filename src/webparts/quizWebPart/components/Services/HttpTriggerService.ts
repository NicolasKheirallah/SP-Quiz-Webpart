// Create a new file: services/HttpTriggerService.ts
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IQuizResult } from '../interfaces'; 

export interface IHttpTriggerPayload {
  userId: string;
  userEmail?: string;
  userName?: string;
  success: boolean;
  scorePercentage: number;
  quizTitle: string;
  resultDate: string;
  siteUrl?: string;
  triggerReason: 'HIGH_SCORE_ACHIEVED';
  threshold: number;
}

export interface IHttpTriggerConfig {
  url: string;
  method: string;
  timeout: number; // This will come from timeLimit property
  includeUserData: boolean;
  customHeaders?: string;
}

export interface IHttpTriggerResponse {
  success: boolean;
  status?: number;
  statusText?: string;
  responseText?: string;
}

export class HttpTriggerService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Sends HTTP trigger when user scores above threshold
   */
  public async sendHighScoreTrigger(
    quizResult: IQuizResult,
    config: IHttpTriggerConfig,
    threshold: number
  ): Promise<boolean> {
    try {
      // Check if score meets threshold
      if (quizResult.ScorePercentage < threshold) {
        console.log(`Score ${quizResult.ScorePercentage}% is below threshold ${threshold}%. HTTP trigger not sent.`);
        return false;
      }

      console.log(`Score ${quizResult.ScorePercentage}% meets threshold ${threshold}%. Sending HTTP trigger...`);

      // Prepare simplified payload
      const payload = this.preparePayload(quizResult, config, threshold);

      // Send HTTP request
      const response = await this.sendHttpRequest(config, payload);

      if (response.success) {
        console.log('HTTP trigger sent successfully');
      } else {
        console.error('Failed to send HTTP trigger');
      }

      return response.success;
    } catch (error) {
      console.error('Error in sendHighScoreTrigger:', error);
      return false;
    }
  }

  /**
   * Prepares the simplified payload for the HTTP trigger
   */
  private preparePayload(
    quizResult: IQuizResult,
    config: IHttpTriggerConfig,
    threshold: number
  ): IHttpTriggerPayload {
    const basePayload: IHttpTriggerPayload = {
      userId: quizResult.UserId,
      success: true, // Always true when this method is called (score met threshold)
      scorePercentage: quizResult.ScorePercentage,
      quizTitle: quizResult.QuizTitle,
      resultDate: quizResult.ResultDate,
      triggerReason: 'HIGH_SCORE_ACHIEVED',
      threshold: threshold,
      siteUrl: this.context.pageContext.web.absoluteUrl
    };

    // Include additional user data if enabled
    if (config.includeUserData) {
      basePayload.userName = quizResult.UserName;
      basePayload.userEmail = quizResult.UserEmail;
    }

    return basePayload;
  }

  /**
   * Sends the HTTP request to the trigger URL
   */
  private async sendHttpRequest(
    config: IHttpTriggerConfig,
    payload: IHttpTriggerPayload
  ): Promise<IHttpTriggerResponse> {
    try {
      // Validate URL
      if (!config.url || !this.isValidUrl(config.url)) {
        console.error('Invalid HTTP trigger URL:', config.url);
        return { success: false, statusText: 'Invalid URL' };
      }

      // Prepare headers
      const headers: HeadersInit = {
        'Content-Type': 'application/json'
      };

      // Add custom headers if provided
      if (config.customHeaders) {
        try {
          const customHeaders = JSON.parse(config.customHeaders);
          Object.assign(headers, customHeaders);
        } catch (parseError) {
          console.warn('Failed to parse custom headers, using defaults:', parseError);
        }
      }

      // Prepare request options
      const requestOptions: RequestInit = {
        method: config.method.toUpperCase(),
        headers: headers,
        body: JSON.stringify(payload)
      };

      // Create AbortController for timeout
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), config.timeout * 1000);

      try {
        console.log('Sending HTTP trigger request:', {
          url: config.url,
          method: config.method,
          payload: payload
        });

        const response = await fetch(config.url, {
          ...requestOptions,
          signal: controller.signal
        });

        clearTimeout(timeoutId);

        const responseText = await response.text();

        if (response.ok) {
          console.log('HTTP trigger request successful:', response.status);
          return { 
            success: true, 
            status: response.status, 
            statusText: response.statusText,
            responseText 
          };
        } else {
          console.error('HTTP trigger request failed:', response.status, response.statusText);
          console.error('Response body:', responseText);
          return { 
            success: false, 
            status: response.status, 
            statusText: response.statusText,
            responseText 
          };
        }
      } catch (fetchError) {
        clearTimeout(timeoutId);
        
        if (fetchError instanceof Error && fetchError.name === 'AbortError') {
          const errorMsg = `HTTP trigger request timed out after ${config.timeout} seconds`;
          console.error(errorMsg);
          return { success: false, statusText: errorMsg };
        } else {
          console.error('HTTP trigger request error:', fetchError);
          return { 
            success: false, 
            statusText: fetchError instanceof Error ? fetchError.message : 'Unknown error' 
          };
        }
      }
    } catch (error) {
      console.error('Error in sendHttpRequest:', error);
      return { 
        success: false, 
        statusText: error instanceof Error ? error.message : 'Unknown error' 
      };
    }
  }

  /**
   * Validates if a string is a valid URL
   */
  private isValidUrl(urlString: string): boolean {
    try {
      const url = new URL(urlString);
      return url.protocol === 'http:' || url.protocol === 'https:';
    } catch {
      return false;
    }
  }
}