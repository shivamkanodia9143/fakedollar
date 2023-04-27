from flask import Flask, request, send_file, render_template
import pandas as pd
import openai

openai.api_key = "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"

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


# Load the ChatGPT model
model = "text-davinci-002"
prompt = "Generate a mapping of elements to activities in retail fitouts."

# Define the Flask app and endpoints
from flask import Flask, request

app = Flask(__name__)


@app.route('/generate_mapping', methods=['POST'])
def generate_mapping_endpoint():
    file = request.files['tasks_file']
    newData = pd.read_excel(file)
    newData2 = newData
    newData2["Keywords"] = ['']*newData.shape[0]
    newData2["Categories"] = ['']*newData.shape[0]
    flag=0
    import json
    for index, row in newData.iterrows():
        prompt_template = "A BOQ element in the construction of retail fitouts contains the name and a description that specifies how to install the element & other requirements. For the following element :-\n"
        prompt_template += f"Element Name: "+ row["Element Name"] + " \n- Description: "+ row["Description"] +"\n"
        # prompt_template += "\nPredict the standard activity name to which this elements belongs to and some keywords :."
        prompt_template += " predict 'category' of work in retail fitouts does this elements description says also provide set of 'keywords' for this:."
        prompt_template += "Answer in python dict with 'keywords' and 'category' as the only keys, remove all keys except 'keywords' and 'category' before answering me, validate your response for json validity before answering me"
        response = openai.Completion.create(
            engine=model,
            prompt=prompt_template,
            max_tokens=2097,
            n=1,
            stop=None,
            temperature=0.5,
        )
        text = response["choices"][0]["text"].replace('\n', '').replace('.','')
        newData2.at[index, "Keywords"] = json.loads(text)["keywords"]
        newData2.at[index, "Category"] = json.loads(text)["category"]
    newData2.to_excel("Cool-gpt-results.xlsx")
    send_file(newData2)


if __name__ == '__main__':
    app.run(debug=True)
