# SharePoint Quiz Web Part - Create SharePoint Lists
# Connect to your SharePoint site first using:
# Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Create QuizQuestions list
Write-Host "Creating QuizQuestions list..." -ForegroundColor Yellow
$quizListExists = Get-PnPList -Identity "QuizQuestions" -ErrorAction SilentlyContinue
if (-not $quizListExists) {
    New-PnPList -Title "QuizQuestions" -Template GenericList
    Add-PnPField -List "QuizQuestions" -DisplayName "Category" -InternalName "Category" -Type Text -AddToDefaultView
    Add-PnPField -List "QuizQuestions" -DisplayName "Choices" -InternalName "Choices" -Type Note -AddToDefaultView
    Write-Host "QuizQuestions list created successfully!" -ForegroundColor Green
} else {
    Write-Host "QuizQuestions list already exists." -ForegroundColor Cyan
}

# Create QuizResults list
Write-Host "Creating QuizResults list..." -ForegroundColor Yellow
$resultsListExists = Get-PnPList -Identity "QuizResults" -ErrorAction SilentlyContinue
if (-not $resultsListExists) {
    New-PnPList -Title "QuizResults" -Template GenericList
    Add-PnPField -List "QuizResults" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -AddToDefaultView
    Add-PnPField -List "QuizResults" -DisplayName "Score" -InternalName "Score" -Type Number -AddToDefaultView
    Add-PnPField -List "QuizResults" -DisplayName "TotalQuestions" -InternalName "TotalQuestions" -Type Number -AddToDefaultView
    Add-PnPField -List "QuizResults" -DisplayName "CompletedDate" -InternalName "CompletedDate" -Type DateTime -AddToDefaultView
    Write-Host "QuizResults list created successfully!" -ForegroundColor Green
} else {
    Write-Host "QuizResults list already exists." -ForegroundColor Cyan
}

Write-Host "SharePoint lists setup complete!" -ForegroundColor Green
