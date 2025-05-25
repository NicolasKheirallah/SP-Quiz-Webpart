# 📝 SharePoint Quiz Web Part

## 🌟 Overview

A feature-rich, interactive quiz solution built for SharePoint Online using the SharePoint Framework (SPFx). This web part enables administrators to create, customize, manage, and deploy quizzes directly within SharePoint sites with a modern user interface and advanced functionality.

## ✨ Features

### 🎯 Quiz Management

- **Comprehensive Question Management**: Create, edit, preview, and delete quiz questions through an intuitive interface
- **Support for multiple question types**:
  - Multiple Choice
  - True/False
  - Multiple Select
  - Short Answer
  - Matching
- **Bulk Operations**: Import/export questions, randomize order, and manage multiple questions at once
- **Category Organization**: Organize questions into categories with customizable ordering

### 🛠️ Advanced Question Configuration

- **Rich Content Support**:
  - Rich text descriptions with formatting
  - Image attachments for questions and answer choices
  - Code snippets with syntax highlighting for programming quizzes
  - Video embedding for multimedia quizzes
- **Customizable Scoring**: Assign custom point values to questions
- **Detailed Explanations**: Add explanations that appear after answering
- **Organization**: Categorize questions for better management with drag-and-drop category ordering
- **Time Limits**: Set per-question time limits or overall quiz duration
- **Case Sensitivity**: Option for case-sensitive answers in short answer questions

### 📊 Quiz Taking Experience

- **Modern UI**: Clean, responsive Fluent UI-based interface
- **Start Page**: Welcome screen with quiz details and instructions
- **Progress Tracking**: Visual indicators of completion progress
- **Save & Resume**: Save progress and continue later with automatic detection
- **Randomization**:
  - Randomize question order
  - Randomize answer choice order
- **Timer Systems**:
  - Question-level timers with visual warnings
  - Overall quiz timer with countdown display
  - Warning notifications as time runs low
- **Pagination**: Configurable questions per page with smooth navigation

### 📈 Results and Reporting

- **Detailed Results**: Comprehensive breakdown of quiz performance with question-by-question analysis
- **Score Visualization**: Graphical representation of scores with performance metrics
- **Answer Review**: Review correct and incorrect answers with explanations
- **Performance Messages**: Customizable feedback based on score ranges
- **SharePoint Integration**: Save results to SharePoint lists for tracking and reporting
- **Custom List Selection**: Choose which SharePoint list to save results to
- **Comprehensive Analytics**: Track answered vs. total questions, percentage correct, and more

### 🔔 HTTP Trigger Integration (NEW!)

- **Webhook Support**: Send HTTP notifications when users achieve high scores
- **Configurable Thresholds**: Set minimum score percentage to trigger notifications
- **Flexible Endpoints**: Support for various HTTP methods (GET, POST, PUT, etc.)
- **Custom Headers**: Add authentication headers or custom metadata
- **User Data Control**: Choose whether to include user information in triggers
- **Simplified Payload**: Lightweight notifications with essential success data only
- **Timeout Configuration**: Configurable request timeouts for reliability

### 🔄 Import/Export Capabilities

- **Multiple Formats**: Support for CSV and JSON formats
- **Template System**: Pre-configured templates for easy question creation
- **Bulk Import**: Add multiple questions at once with validation
- **Data Validation**: Robust validation to ensure data integrity during import

## 🔧 Prerequisites

- SharePoint Online
- Node.js version 16+
- SharePoint Framework (SPFx) 1.13.0 or higher
- Office 365 developer tenant

## 🚀 Installation

### Clone Repository

```bash
git clone https://github.com/NicolasKheirallah/SP-Quiz-Webpart.git
cd SP-Quiz-Webpart
```

### Install Dependencies

```bash
npm install
```

### Build Solution

```bash
gulp bundle --ship
gulp package-solution --ship
```

### Deploy to SharePoint App Catalog

1. Upload the `.sppkg` file from the `sharepoint/solution` folder to your SharePoint App Catalog
2. Deploy the solution globally or to specific sites
3. Add the web part to a SharePoint page

## 📝 Configuration

### Web Part Properties

| Property                    | Description                                     | Default           |
|-----------------------------|-------------------------------------------------|-------------------|
| Title                       | Custom quiz title                               | "SharePoint Quiz" |
| Questions Per Page          | Number of questions per page                    | 5                 |
| Show Progress Indicator     | Display progress tracking                       | True              |
| Randomize Questions         | Shuffle question order                          | False             |
| Randomize Answers           | Shuffle answer choices                          | False             |
| Passing Score               | Minimum percentage required to pass             | 70%               |
| Time Limit                  | Quiz time limit in minutes (0 for unlimited)   | 0 (Unlimited)     |
| Enable Question Time Limit  | Allow per-question time limits                  | False             |
| Default Question Time Limit | Default time limit for questions (seconds)     | 60                |
| Results List Name           | SharePoint list to save quiz results           | "QuizResults"     |
| Enable HTTP Trigger         | Enable webhook notifications for high scores   | False             |
| HTTP Trigger URL            | Webhook endpoint URL                            | ""                |
| HTTP Trigger Score Threshold| Minimum score percentage to trigger webhook    | 80%               |
| HTTP Trigger Method         | HTTP method for webhook requests               | "POST"            |
| HTTP Trigger Timeout        | Request timeout in seconds                      | 30                |
| Include User Data in Trigger| Include user information in webhook payload    | True              |
| Custom Headers              | JSON string of custom headers for webhook      | ""                |
| Custom Messages             | Customizable feedback for different score ranges| Preset messages   |

### HTTP Trigger Configuration

The HTTP trigger sends a simplified payload when users achieve scores above the configured threshold:

```json
{
  "userId": "user@domain.com",
  "userEmail": "user@domain.com",
  "userName": "John Doe",
  "success": true,
  "scorePercentage": 85,
  "quizTitle": "SharePoint Knowledge Test",
  "resultDate": "2023-10-15T14:30:00.000Z",
  "triggerReason": "HIGH_SCORE_ACHIEVED",
  "threshold": 80,
  "siteUrl": "https://tenant.sharepoint.com/sites/sitename"
}
```

**Configuration Examples:**

```json
// Custom Headers Example
{
  "Authorization": "Bearer your-token-here",
  "X-Custom-Header": "CustomValue",
  "Content-Type": "application/json"
}
```

## 🧩 Question Types and Features

### Multiple Choice
- Select one correct answer from multiple options
- Support for images in questions and answer choices
- Point-based scoring
- Optional explanations

### True/False
- Binary choice questions
- Quick knowledge assessment
- Simplified creation

### Multiple Select
- Select multiple correct answers
- Support for partial scoring
- Advanced scenario testing

### Short Answer
- Text-based responses
- Optional case-sensitivity setting
- Pattern matching

### Matching
- Match items in the left column with corresponding items in the right column
- Drag-and-drop interface for intuitive interaction
- Support for multiple pairs in a single question
- Ideal for vocabulary, definitions, and classification questions

### Enhanced Content
- **Rich Text Descriptions**: Format question text with rich editing capabilities
- **Image Support**: Upload and include images in questions and answers
- **Code Snippets**: Include formatted code with language-specific syntax highlighting
- **Video Embedding**: Embed videos from YouTube, Vimeo, Microsoft Stream, or direct URLs
- **Time Limits**: Set question-specific time constraints with visual countdown

## 📤 Import/Export Formats

### JSON Example

```json
[
  {
    "title": "What is SharePoint?",
    "category": "SharePoint Basics",
    "type": "multipleChoice",
    "choices": [
      { "id": "1", "text": "Document Management System", "isCorrect": true },
      { "id": "2", "text": "Programming Language", "isCorrect": false }
    ],
    "points": 2,
    "explanation": "SharePoint is Microsoft's document management and collaboration platform.",
    "timeLimit": 60
  }
]
```

### CSV Example

```csv
Question,Category,Type,Option 1,Option 2,Correct Answer,Points,Explanation,Time Limit
What is SharePoint?,SharePoint Basics,multipleChoice,Document Management,Programming Language,1,2,SharePoint is Microsoft's document management platform,60
```

## 🔄 Save & Resume Functionality

The Quiz Web Part features a robust save and resume capability:

- **Progress Saving**: Save current quiz state including answered questions and remaining time
- **User Association**: Automatically associates saved progress with the current user
- **Resume Notification**: Automatic detection of saved quizzes with option to resume or start fresh
- **Progress Management**: Seamless transition between sessions with complete state preservation
- **Privacy**: Each user's progress is accessible only to them
- **Automatic Cleanup**: Saved progress is automatically removed after successful quiz completion

### SharePoint List for Saving Progress

The web part automatically creates a SharePoint list named "QuizProgress" with the following columns:

| Column Name      | Type                  | Description                                        | Required |
|------------------|------------------------|----------------------------------------------------|----------|
| Title            | Single line of text    | Default column, format: "{Quiz Title} - {User Name} - In Progress" | Yes      |
| UserId           | Single line of text    | User's login name                                  | Yes      |
| UserName         | Single line of text    | User's display name                                | Yes      |
| QuizTitle        | Single line of text    | Title of the quiz                                  | Yes      |
| QuizData         | Multiple lines of text | JSON data of the entire quiz state                 | Yes      |
| LastSaved        | Date and Time          | Timestamp of when progress was saved               | Yes      |

**List Features:**
- **Automatic Creation**: List is created automatically if it doesn't exist
- **User Permissions**: Users can only access their own saved progress
- **Data Validation**: Robust validation ensures data integrity
- **Automatic Cleanup**: Completed quizzes automatically remove saved progress

## 📊 Quiz Results Integration

The web part automatically saves quiz results to a configurable SharePoint list with comprehensive tracking:

- User information (name, email, SharePoint user ID)
- Quiz title and metadata
- Detailed scoring (points, percentage, questions answered)
- Question-by-question breakdown with user responses
- Timestamp and completion data

### SharePoint List for Quiz Results

The web part automatically creates a SharePoint list with your configured name (default: "QuizResults") with the following columns:

| Column Name       | Type                  | Description                                     | Required |
|-------------------|-----------------------|-------------------------------------------------|----------|
| Title             | Single line of text   | Default column, format: "Quiz Result - {Date}"  | Yes      |
| UserName          | Single line of text   | Name of the user who took the quiz              | Yes      |
| UserEmail         | Single line of text   | Email of the user who took the quiz             | No       |
| UserId            | Single line of text   | User's login name                               | Yes      |
| SharePointUserId  | Number                | SharePoint internal user ID                     | No       |
| QuizTitle         | Single line of text   | Title of the quiz                               | Yes      |
| Score             | Number                | Points earned                                   | Yes      |
| TotalPoints       | Number                | Total possible points                           | Yes      |
| ScorePercentage   | Number                | Percentage score                                | Yes      |
| QuestionsAnswered | Number                | Number of questions answered                    | Yes      |
| TotalQuestions    | Number                | Total number of questions                       | Yes      |
| ResultDate        | Date and Time         | When the quiz was submitted                     | Yes      |
| QuestionDetails   | Multiple lines of text| JSON data with detailed question results        | Yes      |

**Enhanced Features:**
- **Automatic List Creation**: Lists are created automatically with all required columns
- **Data Validation**: Comprehensive validation ensures data integrity
- **Rich Analytics**: Detailed question-by-question breakdown for analysis
- **User Association**: Complete user information for tracking and reporting

## 🔒 Permissions

### SharePoint Permissions
- Site Collection Administrator or Site Owner for full functionality
- Edit access to lists for saving results and progress
- Read access for taking quizzes

### API Permissions
- `User.Read` - For user information
- `Sites.ReadWrite.All` - For SharePoint list operations

## 🐛 Troubleshooting

### Common Issues

- **Question Management Issues**: Verify SharePoint permissions for the current user
- **Image Upload Problems**: Check that the FilePicker component has access to the SharePoint context
- **Results Not Saving**: Ensure proper permissions and verify list creation
- **Performance Issues**: Consider reducing the number of questions per page or optimizing images
- **Save & Resume Issues**: Check if QuizProgress list exists and user has proper permissions
- **HTTP Trigger Failures**: Verify URL accessibility and correct configuration

### Error Codes and Solutions

| Error Type | Solution |
|------------|----------|
| List Creation Failed | Check site permissions and ensure user has list creation rights |
| HTTP Trigger Timeout | Increase timeout value or check webhook endpoint availability |
| Image Upload Failed | Verify SharePoint storage limits and file permissions |
| Progress Save Failed | Check QuizProgress list permissions and column configuration |

## 🤝 Contributing

1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add some amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Create Pull Request

### Development Workflow

```bash
# Install dependencies
npm install

# Serve for development
gulp serve

# Build for production
gulp bundle --ship
gulp package-solution --ship

# Run tests
npm test

# Lint code
npm run lint
```

## 🗺️ Roadmap

### Completed ✅
- [x] Matching question type
- [x] Save & Resume functionality
- [x] Video embedding support
- [x] HTTP trigger integration
- [x] Enhanced progress tracking
- [x] Automatic list creation
- [x] Category ordering management

### Planned 🚀
- [ ] Fill-in-the-blank questions
- [ ] Drag-and-drop ordering questions
- [ ] Enhanced analytics and reporting dashboard
- [ ] Quiz templates and sharing
- [ ] Learning path integration
- [ ] Adaptive quizzing based on user performance
- [ ] Multi-language support
- [ ] Mobile app integration
- [ ] Advanced webhook payloads
- [ ] Quiz scheduling and availability windows

## 📄 License

[MIT License](LICENSE)

## 🆘 Support

For issues or feature requests, please [create a GitHub issue](https://github.com/NicolasKheirallah/SP-Quiz-Webpart/issues)

### Getting Help

- **Documentation**: Check this README for common questions
- **Issues**: Search existing issues before creating new ones
- **Discussions**: Use GitHub Discussions for general questions
- **Wiki**: Check the project wiki for detailed guides

## 🙏 Acknowledgments

- Microsoft SharePoint Framework (SPFx)
- Fluent UI React Components
- PnP SPFx Controls and PnP PowerShell
- Prism.js for code syntax highlighting
- uuid for generating unique identifiers
- React Hook Form for form management
- Microsoft Graph API for enhanced SharePoint integration

---

**Made with ❤️ for the SharePoint community**