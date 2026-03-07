import * as React from 'react';
import { IInputs, IOutputs } from '../generated/ManifestTypes';

// Interface for conditional rules
interface ConditionalRule {
  duc_questionpossibleanswersid: string;
  _duc_relatedquestion_value?: string;
  duc_actiontype: number; // 1 = Show, 2 = Hide
}

interface SurveyQuestion {
  duc_surveyquestionid: string;
  duc_label: string;
  duc_required: boolean;
  duc_questiontype: string;
  duc_conditionalrules: ConditionalRule[];
}

export class SurveyForm extends React.Component<IInputs, IOutputs> {
  /*private questions: SurveyQuestion[] = [];
  private conditionalRules: Record<string, ConditionalRule[]> = {};

  constructor(props: IInputs) {
    super(props);
    this.state = {
      answers: new Map(),
    };
  }

  componentDidMount() {
    this.initializeForm();
  }

  componentWillUnmount() {
    this.destroyForm();
  }

  // Initialize the form, populate questions, etc.
  async initializeForm() {
    // Example: Initialize form questions and conditional rules
    await this.loadSurveyQuestions();
    this.applyInitialVisibility();
  }

  // Example function to load survey questions from a service or static data
  async loadSurveyQuestions() {
    // Replace this with an actual API or data call to load questions
    const fetchedQuestions = [
      // Example questions
      {
        duc_surveyquestionid: '1',
        duc_label: 'What is your favorite color?',
        duc_required: true,
        duc_questiontype: 'radio',
        duc_conditionalrules: [],
      },
      {
        duc_surveyquestionid: '2',
        duc_label: 'Why do you like that color?',
        duc_required: true,
        duc_questiontype: 'text',
        duc_conditionalrules: [
          {
            duc_questionpossibleanswersid: '1',
            _duc_relatedquestion_value: '3', // Next question to show based on answer
            duc_actiontype: 1, // Show
          },
        ],
      },
    ];
    
    this.questions = fetchedQuestions;
    this.conditionalRules = this.extractConditionalRules(fetchedQuestions);
  }

  // Extract and format conditional rules
  extractConditionalRules(questions: SurveyQuestion[]) {
    const rules: Record<string, ConditionalRule[]> = {};
    questions.forEach((question) => {
      if (question.duc_conditionalrules) {
        question.duc_conditionalrules.forEach((rule) => {
          if (!rules[question.duc_surveyquestionid]) {
            rules[question.duc_surveyquestionid] = [];
          }
          rules[question.duc_surveyquestionid].push(rule);
        });
      }
    });
    return rules;
  }

  // Apply initial visibility based on conditional rules
  applyInitialVisibility() {
    // Initially hide all conditional questions
    const conditionalQuestionIds = Object.keys(this.conditionalRules);
    conditionalQuestionIds.forEach((questionId) => {
      this.hideQuestion(questionId);
    });
  }

  // Handle the visibility logic for a specific question
  hideQuestion(questionId: string) {
    // Logic to hide the question
    const questionElement = document.querySelector(`[data-question-id="${questionId}"]`);
    if (questionElement) {
      questionElement.classList.add('hidden');
    }
  }

  // Handle form submit
  async submitSurvey() {
    if (this.validateForm()) {
      const answers = this.collectAnswers();
      // Call your submit API here with the answers
    }
  }

  // Collect answers from the form
  collectAnswers() {
    const answers: any[] = [];
    this.questions.forEach((question) => {
      const answer = this.state.answers.get(question.duc_surveyquestionid);
      if (answer) {
        answers.push({
          questionId: question.duc_surveyquestionid,
          value: answer,
        });
      }
    });
    return answers;
  }

  // Validate the form
  validateForm() {
    let isValid = true;
    this.questions.forEach((question) => {
      if (question.duc_required && !this.state.answers.has(question.duc_surveyquestionid)) {
        isValid = false;
        // Show error message or highlight the question
      }
    });
    return isValid;
  }

  render() {
    return (
      <div className="survey-form">
        {this.questions.map((question) => (
          <div key={question.duc_surveyquestionid} data-question-id={question.duc_surveyquestionid} className="question-card">
            <div className="question-label">{question.duc_label}</div>
            {question.duc_questiontype === 'radio' && this.createRadioGroup(question)}
            {question.duc_questiontype === 'text' && this.createTextInput(question)}
            {/* Add more question types as needed *//*}
</div>
))}
<button onClick={() => this.submitSurvey()}>Submit</button>
</div>
);
}

// Render radio button group for a question
createRadioGroup(question: SurveyQuestion) {
return (
<div className="radio-group">
{question.duc_conditionalrules.map((answer, index) => (
<div key={index} className="radio-option">
  <input
    type="radio"
    name={`question_${question.duc_surveyquestionid}`}
    value={answer.duc_questionpossibleanswersid}
    onChange={() => this.handleRadioChange(question, answer)}
  />
  <label>{answer.duc_questionpossibleanswersid}</label>
</div>
))}
</div>
);
}

// Handle radio button change
handleRadioChange(question: SurveyQuestion, answer: ConditionalRule) {
const newAnswers = new Map(this.state.answers);
newAnswers.set(question.duc_surveyquestionid, answer.duc_questionpossibleanswersid);
this.setState({ answers: newAnswers });
}

// Render text input for a question
createTextInput(question: SurveyQuestion) {
return (
<div className="text-input">
<textarea
onChange={(e) => this.handleTextInputChange(question, e.target.value)}
required={question.duc_required}
/>
</div>
);
}

// Handle text input change
handleTextInputChange(question: SurveyQuestion, value: string) {
const newAnswers = new Map(this.state.answers);
newAnswers.set(question.duc_surveyquestionid, value);
this.setState({ answers: newAnswers });
}
*/
}