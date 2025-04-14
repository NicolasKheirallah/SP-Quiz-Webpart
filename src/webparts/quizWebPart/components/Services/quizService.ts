import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Saves quiz results to a SharePoint list.
 * @param spHttpClient - The SPHttpClient instance from the SPFx context.
 * @param webUrl - The absolute URL of the SharePoint site.
 * @param listName - The name of the list where results are saved.
 * @param resultData - A JSON-serializable object containing the quiz result data.
 * @returns A promise resolving to the response JSON.
 */
export async function saveQuizResults(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listName: string,
  resultData: Record<string, unknown>
): Promise<unknown> {
  const endpoint = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`;
  const headers = {
    'Accept': 'application/json;odata=nometadata',
    'Content-type': 'application/json;odata=nometadata',
    'odata-version': ''
  };

  try {
    const response: SPHttpClientResponse = await spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers,
        body: JSON.stringify(resultData)
      }
    );
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to save quiz results: ${errorText}`);
    }
    return await response.json();
  } catch (error) {
    console.error(`Error in saveQuizResults: ${error}`);
    throw error;
  }
}

/**
 * Saves (or updates) quiz progress to a SharePoint list.
 * @param spHttpClient - The SPHttpClient instance from the SPFx context.
 * @param webUrl - The absolute URL of the SharePoint site.
 * @param listName - The name of the list where progress is saved.
 * @param progressData - A JSON-serializable object containing the quiz progress data.
 * @param savedProgressId - (Optional) The existing progress item ID for updates.
 * @returns A promise resolving with the new progress ID if created or void if updated.
 */
export async function saveQuizProgress(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listName: string,
  progressData: Record<string, unknown>,
  savedProgressId?: number
): Promise<number | void> {
  let endpoint: string;
  const method = 'POST';
  const headers: HeadersInit = {
    'Accept': 'application/json;odata=nometadata',
    'Content-type': 'application/json;odata=nometadata',
    'odata-version': ''
  };

  if (savedProgressId) {
    endpoint = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${savedProgressId})`;
    // For updating the existing item, set method override headers
    headers['X-HTTP-Method'] = 'MERGE';
    headers['IF-MATCH'] = '*';
  } else {
    endpoint = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`;
  }

  try {
    const response: SPHttpClientResponse = await spHttpClient.fetch(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        method,
        headers,
        body: JSON.stringify(progressData)
      }
    );
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to save quiz progress: ${errorText}`);
    }
    // If this is a new progress entry, return the new item ID
    if (!savedProgressId) {
      const responseData = await response.json();
      return responseData.Id;
    }
  } catch (error) {
    console.error(`Error in saveQuizProgress: ${error}`);
    throw error;
  }
}

/**
 * Deletes a saved quiz progress item from a SharePoint list.
 * @param spHttpClient - The SPHttpClient instance from the SPFx context.
 * @param webUrl - The absolute URL of the SharePoint site.
 * @param listName - The name of the list where progress is saved.
 * @param progressId - The ID of the progress item to delete.
 * @returns A promise that resolves when the deletion is complete.
 */
export async function deleteSavedProgress(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listName: string,
  progressId: number
): Promise<void> {
  if (!progressId) return;

  const endpoint = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${progressId})`;
  try {
    const response: SPHttpClientResponse = await spHttpClient.fetch(
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
      throw new Error(`Failed to delete saved progress: ${errorText}`);
    }
  } catch (error) {
    console.error(`Error in deleteSavedProgress: ${error}`);
    throw error;
  }
}

/**
 * Checks for existing saved quiz progress for a specific user and quiz.
 * @param spHttpClient - The SPHttpClient instance from the SPFx context.
 * @param webUrl - The absolute URL of the SharePoint site.
 * @param listName - The name of the list where progress is saved.
 * @param userLoginName - The current user’s login name.
 * @param quizTitle - The title of the quiz.
 * @returns A promise resolving with the saved progress item if found, or null otherwise.
 */
export async function checkForSavedProgress(
  spHttpClient: SPHttpClient,
  webUrl: string,
  listName: string,
  userLoginName: string,
  quizTitle: string
): Promise<Record<string, unknown> | undefined> {
  const endpoint = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=UserId eq '${userLoginName}' and QuizTitle eq '${quizTitle}'&$orderby=Modified desc&$top=1`;
  try {
    const response: SPHttpClientResponse = await spHttpClient.get(
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
      throw new Error(`Failed to check for saved progress: ${errorText}`);
    }
    const data = await response.json();
    return data.value && data.value.length > 0 ? data.value[0] : undefined;
  } catch (error) {
    console.error(`Error in checkForSavedProgress: ${error}`);
    throw error;
  }
}
