import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IQuizResult, ISavedQuizProgress } from '../interfaces';

/**
 * Service class for handling quiz data operations with Microsoft Graph API
 * and SharePoint REST API as fallback
 */
export class QuizService {
    private context: WebPartContext;
    private siteId: string;
    private listId: string | null = null;
    private progressListId: string | null = null;

    constructor(context: WebPartContext) {
        this.context = context;
        // Extract site ID from page context
        const absoluteUrl = context.pageContext.web.absoluteUrl;
        const hostname = new URL(absoluteUrl).hostname;
        const sitePath = new URL(absoluteUrl).pathname;
        // Format site ID for Graph API (hostname,spsite,sitePath)
        this.siteId = `${hostname},spsite,${sitePath}`;
    }

    /**
     * Initializes the service by getting list IDs
     * @param resultsListName Name of the SharePoint list for quiz results
     * @param progressListName Name of the SharePoint list for quiz progress
     */
    public async initialize(resultsListName: string, progressListName: string): Promise<void> {
        try {
            // Get list IDs using Graph API
            await this.getListIds(resultsListName, progressListName);
        } catch (error) {
            console.error('Error initializing QuizService:', error);
            throw error;
        }
    }

    /**
     * Gets list IDs using Graph API
     */
    private async getListIds(resultsListName: string, progressListName: string): Promise<void> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Get results list ID
            const resultsResponse = await graphClient
                .api(`/sites/${this.siteId}/lists`)
                .filter(`displayName eq '${resultsListName}'`)
                .get();

            if (resultsResponse.value && resultsResponse.value.length > 0) {
                this.listId = resultsResponse.value[0].id;
            }

            // Get progress list ID
            const progressResponse = await graphClient
                .api(`/sites/${this.siteId}/lists`)
                .filter(`displayName eq '${progressListName}'`)
                .get();

            if (progressResponse.value && progressResponse.value.length > 0) {
                this.progressListId = progressResponse.value[0].id;
            }

            // Create lists if they don't exist
            if (!this.listId) {
                await this.createList(resultsListName, 'Quiz Results');
            }

            if (!this.progressListId) {
                await this.createList(progressListName, 'Quiz Progress');
            }
        } catch (error) {
            console.error('Error getting list IDs:', error);
            // We'll continue without list IDs and fall back to REST API
        }
    }

    /**
     * Creates a new SharePoint list using Graph API
     */
    private async createList(listName: string, listDescription: string): Promise<void> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Create the list
            const response = await graphClient
                .api(`/sites/${this.siteId}/lists`)
                .post({
                    displayName: listName,
                    description: listDescription,
                    list: {
                        template: 'genericList'
                    }
                });

            // Store the list ID
            if (listDescription === 'Quiz Results') {
                this.listId = response.id;
            } else {
                this.progressListId = response.id;
            }

            // Create necessary columns using REST API as Graph doesn't support this easily
            await this.createListColumns(listName);
        } catch (error) {
            console.error(`Error creating list ${listName}:`, error);
            // Fall back to REST API
            await this.createListWithREST(listName, listDescription);
        }
    }

    /**
     * Creates necessary columns for the lists
     */
    private async createListColumns(listName: string): Promise<void> {
        const spHttpClient = this.context.spHttpClient;
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
                await spHttpClient.post(
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
     * Fallback method to create list using REST API
     */
    private async createListWithREST(listName: string, listDescription: string): Promise<void> {
        const spHttpClient = this.context.spHttpClient;
        const webUrl = this.context.pageContext.web.absoluteUrl;

        try {
            const createListResponse = await spHttpClient.post(
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
                throw new Error(`Error creating list with REST API: ${errorText}`);
            }

            // Create columns
            await this.createListColumns(listName);
        } catch (error) {
            console.error(`Error creating list with REST API:`, error);
            throw error;
        }
    }

    /**
     * Saves quiz results using Graph API with REST API fallback
     */
    public async saveQuizResults(quizResult: IQuizResult): Promise<any> {
        try {
            // Try using Graph API first
            if (this.listId) {
                return await this.saveQuizResultsWithGraph(quizResult);
            } else {
                // Fall back to REST API
                return await this.saveQuizResultsWithREST(quizResult);
            }
        } catch (error) {
            console.error('Error in saveQuizResults with Graph API:', error);
            // Fall back to REST API if Graph API fails
            return await this.saveQuizResultsWithREST(quizResult);
        }
    }

    /**
     * Saves quiz results using Graph API
     */
    private async saveQuizResultsWithGraph(quizResult: IQuizResult): Promise<any> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Prepare the item to be saved
            const listItem = {
                fields: {
                    Title: quizResult.Title,
                    UserName: quizResult.UserName,
                    UserId: quizResult.UserId,
                    UserEmail: quizResult.UserEmail,
                    QuizTitle: quizResult.QuizTitle,
                    Score: quizResult.Score,
                    TotalPoints: quizResult.TotalPoints,
                    ScorePercentage: quizResult.ScorePercentage,
                    QuestionsAnswered: quizResult.QuestionsAnswered,
                    TotalQuestions: quizResult.TotalQuestions,
                    QuestionDetails: quizResult.QuestionDetails,
                    ResultDate: quizResult.ResultDate
                }
            };

            // Create the list item
            const response = await graphClient
                .api(`/sites/${this.siteId}/lists/${this.listId}/items`)
                .post(listItem);

            return response;
        } catch (error) {
            console.error('Error saving quiz results with Graph API:', error);
            throw error;
        }
    }

    /**
     * Saves quiz results using REST API
     */
    private async saveQuizResultsWithREST(quizResult: IQuizResult): Promise<any> {
        const spHttpClient = this.context.spHttpClient;
        const webUrl = this.context.pageContext.web.absoluteUrl;
        const resultsListName = 'QuizResults';

        try {
            const response = await spHttpClient.post(
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

            return await response.json();
        } catch (error) {
            console.error('Error saving quiz results with REST API:', error);
            throw error;
        }
    }

    /**
     * Saves quiz progress using Graph API with REST API fallback
     */
    public async saveQuizProgress(
        progressData: ISavedQuizProgress,
        savedProgressId?: number
    ): Promise<number | void> {
        try {
            // Try using Graph API first
            if (this.progressListId) {
                return await this.saveQuizProgressWithGraph(progressData, savedProgressId);
            } else {
                // Fall back to REST API
                return await this.saveQuizProgressWithREST(progressData, savedProgressId);
            }
        } catch (error) {
            console.error('Error saving quiz progress with Graph API:', error);
            // Fall back to REST API if Graph API fails
            return await this.saveQuizProgressWithREST(progressData, savedProgressId);
        }
    }

    /**
     * Saves quiz progress using Graph API
     */
    private async saveQuizProgressWithGraph(
        progressData: ISavedQuizProgress,
        savedProgressId?: number
    ): Promise<number | void> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Prepare the item to be saved
            const listItem = {
                fields: {
                    Title: `${progressData.quizTitle} - ${progressData.userName} - In Progress`,
                    UserId: progressData.userId,
                    UserName: progressData.userName,
                    QuizTitle: progressData.quizTitle,
                    QuizData: JSON.stringify(progressData),
                    LastSaved: progressData.lastSaved
                }
            };

            if (savedProgressId) {
                // Update existing item
                await graphClient
                    .api(`/sites/${this.siteId}/lists/${this.progressListId}/items/${savedProgressId}`)
                    .update(listItem);

                return savedProgressId;
            } else {
                // Create new item
                const response = await graphClient
                    .api(`/sites/${this.siteId}/lists/${this.progressListId}/items`)
                    .post(listItem);

                return response.id;
            }
        } catch (error) {
            console.error('Error saving quiz progress with Graph API:', error);
            throw error;
        }
    }

    /**
     * Saves quiz progress using REST API
     */
    private async saveQuizProgressWithREST(
        progressData: ISavedQuizProgress,
        savedProgressId?: number
    ): Promise<number | void> {
        const spHttpClient = this.context.spHttpClient;
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
            const response: SPHttpClientResponse = await spHttpClient.fetch(
                endpoint,
                SPHttpClient.configurations.v1,
                {
                    method,
                    headers,
                    body: JSON.stringify(spItemData)
                }
            );

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Failed to save progress: ${JSON.stringify(errorData)}`);
            }

            // If this is a new record, get the ID
            if (!savedProgressId) {
                const responseData = await response.json();
                return responseData.Id;
            }
        } catch (error) {
            console.error('Error saving quiz progress with REST API:', error);
            throw error;
        }
    }

    /**
     * Gets saved quiz progress using Graph API with REST API fallback
     */
    public async getSavedProgress(
        userLoginName: string,
        quizTitle: string
    ): Promise<ISavedQuizProgress | undefined> {
        try {
            // Try using Graph API first
            if (this.progressListId) {
                return await this.getSavedProgressWithGraph(userLoginName, quizTitle);
            } else {
                // Fall back to REST API
                return await this.getSavedProgressWithREST(userLoginName, quizTitle);
            }
        } catch (error) {
            console.error('Error getting saved progress with Graph API:', error);
            // Fall back to REST API if Graph API fails
            return await this.getSavedProgressWithREST(userLoginName, quizTitle);
        }
    }

    /**
     * Gets saved quiz progress using Graph API
     */
    private async getSavedProgressWithGraph(
        userLoginName: string,
        quizTitle: string
    ): Promise<ISavedQuizProgress | undefined> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Query for saved progress
            const response = await graphClient
                .api(`/sites/${this.siteId}/lists/${this.progressListId}/items`)
                .expand('fields')
                .filter(`fields/UserId eq '${userLoginName}' and fields/QuizTitle eq '${quizTitle}'`)
                .orderby('fields/LastSaved desc')
                .top(1)
                .get();

            if (response.value && response.value.length > 0) {
                const savedItem = response.value[0];

                // Parse the saved progress data
                try {
                    const progressData: ISavedQuizProgress = JSON.parse(savedItem.fields.QuizData);
                    // Add the item ID so we can update it later
                    progressData.id = savedItem.id;
                    return progressData;
                } catch (parseError) {
                    console.error('Error parsing saved progress data:', parseError);
                    return undefined;
                }
            }

            return undefined;
        } catch (error) {
            console.error('Error getting saved progress with Graph API:', error);
            throw error;
        }
    }

    /**
     * Gets saved quiz progress using REST API
     */
    private async getSavedProgressWithREST(
        userLoginName: string,
        quizTitle: string
    ): Promise<ISavedQuizProgress | undefined> {
        const spHttpClient = this.context.spHttpClient;
        const webUrl = this.context.pageContext.web.absoluteUrl;
        const progressListName = 'QuizProgress';

        const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items?$filter=UserId eq '${userLoginName}' and QuizTitle eq '${quizTitle}'&$orderby=LastSaved desc&$top=1`;

        try {
            const response = await spHttpClient.get(
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

            const data = await response.json();

            if (data.value && data.value.length > 0) {
                const savedItem = data.value[0];

                // Parse the saved progress data
                try {
                    const progressData: ISavedQuizProgress = JSON.parse(savedItem.QuizData);
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
            console.error('Error getting saved progress with REST API:', error);
            throw error;
        }
    }

    /**
     * Deletes saved quiz progress using Graph API with REST API fallback
     */
    public async deleteSavedProgress(progressId: number): Promise<void> {
        try {
            // Try using Graph API first
            if (this.progressListId) {
                await this.deleteSavedProgressWithGraph(progressId);
            } else {
                // Fall back to REST API
                await this.deleteSavedProgressWithREST(progressId);
            }
        } catch (error) {
            console.error('Error deleting saved progress with Graph API:', error);
            // Fall back to REST API if Graph API fails
            await this.deleteSavedProgressWithREST(progressId);
        }
    }

    /**
     * Deletes saved quiz progress using Graph API
     */
    private async deleteSavedProgressWithGraph(progressId: number): Promise<void> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Delete the item
            await graphClient
                .api(`/sites/${this.siteId}/lists/${this.progressListId}/items/${progressId}`)
                .delete();
        } catch (error) {
            console.error('Error deleting saved progress with Graph API:', error);
            throw error;
        }
    }

    /**
     * Deletes saved quiz progress using REST API
     */
    private async deleteSavedProgressWithREST(progressId: number): Promise<void> {
        const spHttpClient = this.context.spHttpClient;
        const webUrl = this.context.pageContext.web.absoluteUrl;
        const progressListName = 'QuizProgress';

        const endpoint = `${webUrl}/_api/web/lists/getbytitle('${progressListName}')/items(${progressId})`;

        try {
            const response = await spHttpClient.fetch(
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
            console.error('Error deleting saved progress with REST API:', error);
            throw error;
        }
    }

    /**
     * Performs batch operations using Microsoft Graph API to improve performance
     * @param requests Array of request objects for batch processing
     * @returns Array of responses from the batch operation
     */
    private async performBatchOperations(requests: {
        id: string;
        method: string;
        url: string;
        body?: any;
        headers?: Record<string, string>;
    }[]): Promise<any[]> {
        try {
            const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

            // Format the batch request
            const batchRequestBody = {
                requests: requests.map(request => ({
                    id: request.id,
                    method: request.method,
                    url: request.url,
                    body: request.body,
                    headers: request.headers || {}
                }))
            };

            // Send the batch request
            const batchResponse = await graphClient
                .api('/$batch')
                .post(batchRequestBody);

            // Process and return the responses
            if (batchResponse && batchResponse.responses) {
                return batchResponse.responses;
            }

            return [];
        } catch (error) {
            console.error('Error performing batch operations:', error);
            throw error;
        }
    }

    /**
     * Batch operation for bulk quiz result submission
     * @param quizResults Array of quiz results to save
     * @returns Array of responses from the batch operation
     */
    public async bulkSaveQuizResults(quizResults: IQuizResult[]): Promise<any[]> {
        try {
            if (!this.listId || quizResults.length === 0) {
                throw new Error('List ID not available or no results to save');
            }

            // Prepare batch requests
            const batchRequests = quizResults.map((result, index) => {
                return {
                    id: `result-${index}`,
                    method: 'POST',
                    url: `/sites/${this.siteId}/lists/${this.listId}/items`,
                    body: {
                        fields: {
                            Title: result.Title,
                            UserName: result.UserName,
                            UserId: result.UserId,
                            UserEmail: result.UserEmail,
                            QuizTitle: result.QuizTitle,
                            Score: result.Score,
                            TotalPoints: result.TotalPoints,
                            ScorePercentage: result.ScorePercentage,
                            QuestionsAnswered: result.QuestionsAnswered,
                            TotalQuestions: result.TotalQuestions,
                            QuestionDetails: result.QuestionDetails,
                            ResultDate: result.ResultDate
                        }
                    }
                };
            });

            // Execute batch operation
            return await this.performBatchOperations(batchRequests);
        } catch (error) {
            console.error('Error in bulkSaveQuizResults:', error);

            // Fall back to individual saves
            const results = [];
            for (const quizResult of quizResults) {
                try {
                    const result = await this.saveQuizResults(quizResult);
                    results.push({ status: 200, id: quizResult.Title, body: result });
                } catch (saveError) {
                    results.push({
                        status: 500,
                        id: quizResult.Title,
                        body: {
                            error: {
                                message: saveError instanceof Error
                                    ? saveError.message
                                    : 'Unknown error occurred'
                            }
                        }
                    });
                }
            }

            return results;
        }
    }

    /**
     * Batch operation for bulk quiz progress checking
     * @param userLoginName User login name
     * @param quizTitles Array of quiz titles to check progress for
     * @returns Record of quiz titles to saved progress data
     */
    public async bulkCheckSavedProgress(
        userLoginName: string,
        quizTitles: string[]
    ): Promise<Record<string, ISavedQuizProgress | undefined>> {
        try {
            if (!this.progressListId || quizTitles.length === 0) {
                throw new Error('Progress list ID not available or no quiz titles provided');
            }

            // Prepare batch requests
            const batchRequests = quizTitles.map((title, index) => {
                return {
                    id: `progress-${index}`,
                    method: 'GET',
                    url: `/sites/${this.siteId}/lists/${this.progressListId}/items?$expand=fields&$filter=fields/UserId eq '${userLoginName}' and fields/QuizTitle eq '${title}'&$orderby=fields/LastSaved desc&$top=1`
                };
            });

            // Execute batch operation
            const batchResponses = await this.performBatchOperations(batchRequests);

            // Process responses
            const progressMap: Record<string, ISavedQuizProgress | undefined> = {};

            batchResponses.forEach((response, index) => {
                const quizTitle = quizTitles[index];

                if (response.status === 200 && response.body.value && response.body.value.length > 0) {
                    const savedItem = response.body.value[0];

                    try {
                        const progressData: ISavedQuizProgress = JSON.parse(savedItem.fields.QuizData);
                        progressData.id = savedItem.id;
                        progressMap[quizTitle] = progressData;
                    } catch (parseError) {
                        console.error(`Error parsing saved progress data for ${quizTitle}:`, parseError);
                        progressMap[quizTitle] = undefined;
                    }
                } else {
                    progressMap[quizTitle] = undefined;
                }
            });

            return progressMap;
        } catch (error) {
            console.error('Error in bulkCheckSavedProgress:', error);

            // Fall back to individual checks
            const progressMap: Record<string, ISavedQuizProgress | undefined> = {};

            for (const quizTitle of quizTitles) {
                try {
                    progressMap[quizTitle] = await this.getSavedProgress(userLoginName, quizTitle);
                } catch (checkError) {
                    console.error(`Error checking progress for ${quizTitle}:`, checkError);
                    progressMap[quizTitle] = undefined;
                }
            }

            return progressMap;
        }
    }
}