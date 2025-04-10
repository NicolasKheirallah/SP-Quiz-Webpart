# 📝 SharePoint Quiz Web Part

## 🌟 Overview

A flexible and interactive quiz solution built for SharePoint Online using the SharePoint Framework (SPFx), enabling administrators to create, manage, and deploy quizzes directly within SharePoint sites.

## ✨ Features

### 🎯 Quiz Management

- Create, edit, and delete quiz questions
- Support for multiple question types:
  - Multiple Choice
  - True/False
  - Multiple Select
  - Short Answer

### 🛠 Question Configuration

- Assign points to questions
- Add explanations
- Categorize questions
- Import/Export questions via CSV or JSON

### 🖥 Quiz Taking Experience

- Randomize questions and answer choices
- Progress tracking
- Time limit support
- Detailed results reporting
- Responsive design

## 🔧 Prerequisites

- SharePoint Online
- Node.js version 16+
- SharePoint Framework (SPFx)
- Office 365 developer tenant

## 🚀 Installation

### Clone Repository

```bash
git clone https://your-repository-url.git
cd sp-quiz-webpart
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

## 📝 Configuration

### Web Part Properties

| Property            | Description                  | Default           |
| ------------------- | ---------------------------- | ----------------- |
| Title               | Custom quiz title            | "SharePoint Quiz" |
| Questions Per Page  | Number of questions per page | 5                 |
| Randomize Questions | Shuffle question order       | False             |
| Randomize Answers   | Shuffle answer choices       | False             |
| Passing Score       | Minimum score to pass        | 70%               |
| Time Limit          | Optional quiz duration       | Unlimited         |

## 🧩 Question Types

### Multiple Choice

- Select one correct answer
- Point-based scoring
- Optional explanations

### True/False

- Binary choice questions
- Quick knowledge assessment

### Multiple Select

- Select multiple correct answers
- Complex evaluation scenarios

### Short Answer

- Text-based responses
- Optional case-sensitivity

## 📤 Import/Export Formats

### JSON Example

```json
[
  {
    "title": "What is SharePoint?",
    "category": "SharePoint Basics",
    "type": "multipleChoice",
    "choices": [
      { "text": "Document Management System", "isCorrect": true },
      { "text": "Programming Language", "isCorrect": false }
    ]
  }
]
```

### CSV Example

```csv
Question,Category,Type,Option 1,Option 2,Correct Answer
What is SharePoint?,SharePoint Basics,multipleChoice,Document Management,Programming Language,1
```

## 🔒 Permissions

### SharePoint Permissions

- Site Collection Administrator
- Site Owner
- Edit access to lists

### API Permissions

- `User.Read`
- `Sites.ReadWrite.All`

## 🐛 Troubleshooting

- Verify SharePoint permissions
- Check browser console
- Confirm SharePoint Framework version compatibility

## 🤝 Contributing

1. Fork the repository
2. Create feature branch
3. Commit changes
4. Push to branch
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

## 🗺 Roadmap

- [ ] Additional question types
- [ ] Enhanced reporting
- [ ] Learning Management System integration
- [ ] Adaptive quizzing

## 📄 License

[Specify your license, e.g., MIT]

## 🆘 Support

For issues or feature requests, please [create a GitHub issue](https://github.com/your-repo/issues)

## 🙏 Acknowledgments

- Microsoft SharePoint Framework
- Fluent UI
- PnP SPFx Controls

```

This Markdown README provides a comprehensive, well-structured guide with emojis for visual appeal and clear sections covering installation, configuration, features, and contribution guidelines.

Would you like me to elaborate on any section or adjust the formatting?
```
