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
- **Bulk Operations**: Import/export questions, randomize order, and manage multiple questions at once

### 🛠️ Advanced Question Configuration

- **Rich Content Support**:
  - Rich text descriptions with formatting
  - Image attachments for questions and answer choices
  - Code snippets with syntax highlighting for programming quizzes
- **Customizable Scoring**: Assign custom point values to questions
- **Detailed Explanations**: Add explanations that appear after answering
- **Organization**: Categorize questions for better management
- **Time Limits**: Set per-question time limits or overall quiz duration
- **Case Sensitivity**: Option for case-sensitive answers in short answer questions

### 📊 Quiz Taking Experience

- **Modern UI**: Clean, responsive Fluent UI-based interface
- **Start Page**: Welcome screen with quiz details and instructions
- **Progress Tracking**: Visual indicators of completion progress
- **Randomization**:
  - Randomize question order
  - Randomize answer choice order
- **Timer Systems**:
  - Question-level timers
  - Overall quiz timer with visual indicators
  - Warning notifications as time runs low
- **Pagination**: Configurable questions per page

### 📈 Results and Reporting

- **Detailed Results**: Comprehensive breakdown of quiz performance
- **Score Visualization**: Graphical representation of scores
- **Answer Review**: Review correct and incorrect answers with explanations
- **Performance Messages**: Customizable feedback based on score ranges
- **SharePoint Integration**: Save results to SharePoint lists for tracking

### 🔄 Import/Export Capabilities

- **Multiple Formats**: Support for CSV and JSON formats
- **Template System**: Pre-configured templates for easy question creation
- **Bulk Import**: Add multiple questions at once
- **Validation**: Robust validation to ensure data integrity

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

| Property               | Description                                     | Default           |
|------------------------|-------------------------------------------------|-------------------|
| Title                  | Custom quiz title                               | "SharePoint Quiz" |
| Questions Per Page     | Number of questions per page                    | 5                 |
| Show Progress Indicator| Display progress tracking                       | True              |
| Randomize Questions    | Shuffle question order                          | False             |
| Randomize Answers      | Shuffle answer choices                          | False             |
| Passing Score          | Minimum percentage required to pass             | 70%               |
| Time Limit             | Quiz time limit in minutes (0 for unlimited)    | 0 (Unlimited)     |
| Custom Messages        | Customizable feedback for different score ranges| Preset messages   |

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

### Enhanced Content
- **Rich Text Descriptions**: Format question text with rich editing capabilities
- **Image Support**: Upload and include images in questions and answers
- **Code Snippets**: Include formatted code with language-specific syntax highlighting
- **Time Limits**: Set question-specific time constraints

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
    "explanation": "SharePoint is Microsoft's document management and collaboration platform."
  }
]
```

### CSV Example

```csv
Question,Category,Type,Option 1,Option 2,Correct Answer,Points,Explanation
What is SharePoint?,SharePoint Basics,multipleChoice,Document Management,Programming Language,1,2,SharePoint is Microsoft's document management platform
```

## 📊 Quiz Results Integration

The web part automatically saves quiz results to a SharePoint list called "QuizResults" with the following data:

- User information (name, email)
- Quiz title
- Score details (points, percentage)
- Question-by-question breakdown
- Timestamp

## 🔒 Permissions

### SharePoint Permissions
- Site Collection Administrator or Site Owner for full functionality
- Edit access to lists for saving results

### API Permissions
- `User.Read`
- `Sites.ReadWrite.All`

## 🐛 Troubleshooting

- **Question Management Issues**: Verify SharePoint permissions for the current user
- **Image Upload Problems**: Check that the FilePicker component has access to the SharePoint context
- **Results Not Saving**: Ensure the QuizResults list exists and the user has proper permissions
- **Performance Issues**: Consider reducing the number of questions per page or images in quizzes

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
```

## 🗺️ Roadmap

- [ ] Matching question type
- [ ] Fill-in-the-blank questions
- [ ] Drag-and-drop ordering questions
- [ ] Enhanced analytics and reporting dashboard
- [ ] Quiz sharing and collaboration features
- [ ] Learning path integration
- [ ] Adaptive quizzing based on user performance
- [ ] Multi-language support

## 📄 License

[MIT License](LICENSE)

## 🆘 Support

For issues or feature requests, please [create a GitHub issue](https://github.com/NicolasKheirallah/SP-Quiz-Webpart/issues)

## 🙏 Acknowledgments

- Microsoft SharePoint Framework (SPFx)
- Fluent UI React Components
- PnP SPFx Controls
- Prism.js for code syntax highlighting
- uuid for generating unique identifiers