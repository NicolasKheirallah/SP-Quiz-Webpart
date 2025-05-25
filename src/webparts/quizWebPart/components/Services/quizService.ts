import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IQuizResult, ISavedQuizProgress } from '../interfaces';

// Define interfaces for SharePoint API responses
interface ISharePointListResponse {
  value: ISharePointItem[];
}

interface ISharePointItem {
  Id: number;
  Title?: string;
  QuizData?: string;
  [key: string]: unknown;
}

interface ISharePointItemResponse {
  Id: number;
  [key: string]: unknown;
}

/**
 * Service class for handling quiz data operations with SharePoint REST API
 */
export class QuizService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Ensures the required lists exist
   */
  public async ensureLists(resultsListName: string, progressListName: string): Promise<void> {
    // Check if results list exists, create if not
    const resultsListExists = await this.checkIfListExists(resultsListName);
    if (!resultsListExists) {
      await this.createList(resultsListName, 'Stores quiz results for the Quiz Web Part');
    }

    // Check if progress list exists, create if not
    const progressListExists = await this.checkIfListExists(progressListName);
    if (!progressListExists) {
      await this.createList(progressListName, 'Stores in-progress quiz data for the Quiz Web Part');
    }
  }

  /**
   * Checks if a list exists
   */
  private async checkIfListExists(listName: string): Promise<boolean> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const endpoint = `${webUrl}/_api/web/lists/getbytitle('${listName}')`;
      
      const response = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );
      
      return response.ok;
    } catch (error) {
      console.error(`Error checking if list '${listName}' exists:`, error);
      return false;
    }
  }

  /**
   * Creates a new SharePoint list
   */
  private async createList(listName: string, listDescription: string): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      
      // Create the list
      const createListResponse = await this.context.spHttpClient.post(
        `${webUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify({
            Title: listName,
            BaseTemplate: 100, // Custom list
            ContentTypesEnabled: false,
            Description: listDescription
          })
        }
      );
      
      if (!createListResponse.ok) {
        const errorText = await createListResponse.text();
        throw new Error(`Error creating list: ${errorText}`);
      }
      
      // Create necessary columns
      await this.createListColumns(listName);
    } catch (error) {
      console.error(`Error creating list ${listName}:`, error);
      throw error;
    }
  }
  
  /**
   * Creates necessary columns for the lists
   */
  private async createListColumns(listName: string): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    
    const columns = [
      { Title: "UserName", FieldTypeKind: 2 },      // Text field
      { Title: "UserId", FieldTypeKind: 2 },        // Text field
      { Title: "UserEmail", FieldTypeKind: 2 },     // Text field
      { Title: "QuizTitle", FieldTypeKind: 2 },     // Text field
      { Title: "Score", FieldTypeKind: 9 },         // Number field
      { Title: "TotalPoints", FieldTypeKind: 9 },   // Number field
      { Title: "ScorePercentage", FieldTypeKind: 9 }, // Number field
      { Title: "QuestionsAnswered", FieldTypeKind: 9 }, // Number field
      { Title: "TotalQuestions", FieldTypeKind: 9 },    // Number field
      { Title: "QuestionDetails", FieldTypeKind: 3 },   // Multi-line text field
      { Title: "ResultDate", FieldTypeKind: 4 },        // DateTime field
      { Title: "QuizData", FieldTypeKind: 3 },          // Multi-line text field (for progress)
      { Title: "LastSaved", FieldTypeKind: 4 }          // DateTime field (for progress)
    ];

    for (const column of columns) {
      try {
        await this.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'odata-version': ''
            },
            body: JSON.stringify(column)
          }
        );
      } catch (error) {
        console.error(`Error creating column ${column.Title}:`, error);
        // Continue with other columns even if one fails
      }
    }
  }
  
  /**
   * Saves quiz results to a SharePoint list.
   * @param resultData A JSON-serializable object containing the quiz result data.
   * @returns A promise resolving to the response JSON.
   */
  public async saveQuizResults(quizResult: IQuizResult): Promise<ISharePointItemResponse> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const resultsListName = 'QuizResults';
    
    try {
      const response = await this.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${resultsListName}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify(quizResult)
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to save quiz results: ${errorText}`);
      }
      
      return await response.json() as ISharePointItemResponse;
    } catch (error) {
      console.error(`Error saving quiz results:`, error);
      throw error;
    }
  }
  
  /**
   * Saves quiz progress to a SharePoint list.
   * @param progressData A JSON-serializable object containing the quiz progress data.
   * @param savedProgressId (Optional) The existing progress item ID for updates.
   * @returns A promise resolving with the new progress ID if created or void if updated.
   */
  public async saveQuizProgress(
    progressData: ISavedQuizProgress,
    savedProgressId?: number
  ): Promise<number | void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const progressListName = 'QuizProgress';
    
    let endpoint: string;
    const method = 'POST';
    const headers: HeadersInit = {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': ''
    };
    
    // Convert to SharePoint item format
    const spItemData = {
      Title: `${progressData.quizTitle} - ${progressData.userName} - In Progress`,
      UserId: progressData.userId,
      UserName: progressData.userName,
      QuizTitle: progressData.quizTitle,
      QuizData: JSON.stringify(progressData),
      LastSaved: progressData.lastSaved
    };
    
    if (savedProgressId) {
      // Update existing record
      endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${savedProgressId})`;
      headers['X-HTTP-Method'] = 'MERGE';
      headers['IF-MATCH'] = '*';
    } else {
      // Create new record
      endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items`;
    }
    
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.fetch(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          method,
          headers,
          body: JSON.stringify(spItemData)
        }
      );
      
      if (!response.ok) {
        const errorData = await response.json() as Record<string, unknown>;
        throw new Error(`Failed to save progress: ${JSON.stringify(errorData)}`);
      }
      
      // If this is a new record, get the ID
      if (!savedProgressId) {
        const responseData = await response.json() as ISharePointItemResponse;
        return responseData.Id;
      }
    } catch (error) {
      console.error('Error saving quiz progress:', error);
      throw error;
    }
  }
  
  /**
   * Gets saved quiz progress for a specific user and quiz
   * @param userLoginName The user's login name
   * @param quizTitle The title of the quiz
   * @returns A promise resolving with the saved progress if found, or undefined
   */
  public async getSavedProgress(
    userLoginName: string,
    quizTitle: string
  ): Promise<ISavedQuizProgress | undefined> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const progressListName = 'QuizProgress';
    
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items?$filter=UserId eq '${userLoginName}' and QuizTitle eq '${quizTitle}'&$orderby=LastSaved desc&$top=1`;
    
    try {
      const response = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Error response when retrieving saved progress: ${errorText}`);
      }
      
      const data = await response.json() as ISharePointListResponse;
      
      if (data.value && data.value.length > 0) {
        const savedItem = data.value[0];
        
        // Parse the saved progress data
        try {
          const progressData: ISavedQuizProgress = JSON.parse(savedItem.QuizData || '{}');
          // Add the item ID so we can update it later
          progressData.id = savedItem.Id;
          return progressData;
        } catch (parseError) {
          console.error('Error parsing saved progress data:', parseError);
          return undefined;
        }
      }
      
      return undefined;
    } catch (error) {
      console.error('Error getting saved progress:', error);
      throw error;
    }
  }
  
  /**
   * Deletes a saved quiz progress item
   * @param progressId The ID of the progress item to delete
   */
  public async deleteSavedProgress(progressId: number): Promise<void> {
    if (!progressId) return;
    
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const progressListName = 'QuizProgress';
    
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${progressId})`;
    
    try {
      const response = await this.context.spHttpClient.fetch(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          method: 'DELETE',
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Error deleting saved progress: ${errorText}`);
      }
    } catch (error) {
      console.error('Error deleting saved progress:', error);
      throw error;
    }
  }
  
  /**
   * Processes multiple quiz operations in sequence
   * @param operations Array of functions that return Promises
   * @returns Promise that resolves when all operations are complete
   */
  public async processBatchOperations<T>(
    operations: (() => Promise<T>)[]
  ): Promise<T[]> {
    const results: T[] = [];
    
    for (const operation of operations) {
      try {
        const result = await operation();
        results.push(result);
      } catch (error) {
        console.error('Error in batch operation:', error);
        throw error;
      }
    }
    
    return results;
  }
  
  /**
   * Gets quiz results for a specific user
   * @param userLoginName The user's login name
   * @returns Promise resolving to the user's quiz results
   */
  public async getQuizResults(userLoginName: string): Promise<ISharePointItem[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const resultsListName = 'QuizResults';
    
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('${resultsListName}')/items?$filter=UserId eq '${userLoginName}'&$orderby=ResultDate desc`;
    
    try {
      const response = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }
      );
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Error retrieving quiz results: ${errorText}`);
      }
      
      const data = await response.json() as ISharePointListResponse;
      return data.value || [];
    } catch (error) {
      console.error('Error getting quiz results:', error);
      throw error;
    }
  }
}