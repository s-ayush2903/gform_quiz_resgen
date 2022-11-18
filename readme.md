# gform_quiz_resgen

Generates results from the quizzes / tests conducted on Google Forms. Requires a CSV file containing entries of students appeared with the answers filled by them corresponding to each question.

## Running the Code
Given instructions should be followed in the order which they are given below:  
* Create a new [venv](https://docs.python.org/3/library/venv.html) via `python3 -m venv ./cenv` (`cenv` can be replaced by _any_ name which you want, it is the name of the virtual env you want to create).
* Activate the venv via `source cenv/bin/activate`
* Install dependencies via `pip3 install -r rex`
* Start the server via `python3 main.py`
* `^Z` / `^C` to stop / kill the server
* `deactivate` to exit from created venv
