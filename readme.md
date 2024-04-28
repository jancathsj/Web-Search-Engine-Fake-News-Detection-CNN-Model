WEB SEARCH ENGINE FAKE NEWS DETECTION USING TEXTUAL DATA AND URL REPRESENTATION THROUGH CNN
Author: Jan Catherine S. San Juan
Contact: jancatherinesanjuan@gmail.com, jssanjuan1@up.edu.ph

This project uses CNN and baseline traditional machine learning algorithms to classify WSE results as either Fake News or Real News. 


I. CNN Models
Prerequisite:
Tensorflow: Run "pip install tensorflow"

Input:
For the CNN, the format of the input must be a csv file with "Title", "Decription", "URL", and "Fake" columns. The Fake column must have a value of 0 for Real News and 1 for Fake News.

1. To run all CNN models, open the CNN Model folder and open "Concat.ipynb" on JupyterLab.
2. Change the directory in the third line to where your test dataset will be:
os.chdir('../Datasets') 
2. Change the input file for the "testdata" to the dataset you wish the models to run.
3. Run the file and the results for all models will be presented.


II. Machine learning
Input:
The input dataset for this section must first go through Feature extraction through TF-IDF.ipynb and manual feature extraction through the modified version of Jain, Bhaskar, Srikanth, and Ramakrishnan's code in the folder "url_classification_dl-main" [1]. The results must then be standardized through the KNNImputer.ipynb and MinMaxScalar.ipynb. The feature selection is done through Chi-Square-01.ipynb and Pearson_correlation.ipynb. The final input must be a csv file with the unique words and extracted URL features as columns.

1. To run all Machine learning models, open the MachineLearning Model folder and open "K-Fold-02.ipynb" on JupyterLab.
2. Change the input file for the "testdata" to the dataset you wish the models to run.
3. Run the file and the results for all models will be presented.


III. Preprocessing folder
This includes the file to obtain the list of foreign words and the word count. This also includes the file to clean the dataset from certain punctuation marks, special characters, foreign words, and stop words.



[1]A. Jain, A. Bhaskar, Srikanth, and R. Ramakrishnan, “Url Feature Extraction & Classification,” Github, Dec. 11, 2021. https://github.com/Rohith-2/url_classification_dl
