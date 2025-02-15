from wtforms import Form, StringField, RadioField, SelectField, TextAreaField, validators, IntegerField, EmailField
from wtforms.fields import EmailField, DateField


class CreateItemForm(Form):
    category = SelectField('Category', [validators.DataRequired()],
                           choices=[('', 'Select'), ('Newspapers/Paper', 'Newspaper/Paper'), ('Cardboard', 'Cardboard'),
                                    ('Clothing', 'Clothing'), ('Metal Recyclables', 'Metal Recyclables'), ('Furniture', 'Furniture'),
                                    ('Electronics', 'Electronics'), ('Speakers', 'Speakers')], default='')
    item = StringField('Item', [validators.Length(min=1, max=100), validators.DataRequired()])
    description = StringField('Description', [validators.Length(min=1, max=150), validators.DataRequired()])
    condition = SelectField('Condition', [validators.DataRequired()],
                            choices=[('', 'Select'), ('Good', 'Good'), ('Bad', 'Bad')], default='')
    stock = IntegerField('Stock', [validators.DataRequired()])
    selling_price = IntegerField('Selling Price ($)', [validators.DataRequired()])


class CreateFeedbackForm(Form):
    name = StringField('Name', [validators.Length(min=1, max=100), validators.DataRequired()])
    email = EmailField('Email', [validators.DataRequired()])
    feedback = TextAreaField('Feedback', [validators.DataRequired()])


class CreateReplyFeedbackForm(Form):
    subject = StringField('Subject', [validators.Length(min=1, max=100), validators.DataRequired()])
    recipient_email = EmailField('Recipient Email', [validators.DataRequired()])
    message = TextAreaField('Message', [validators.DataRequired()])
