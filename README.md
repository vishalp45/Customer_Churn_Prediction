# Customer Churn Prediction

## Project Overview

Customer churn is one of the most important business problems in the telecom industry. This project aims to predict whether a customer is likely to leave the service using machine learning techniques and identify the key factors influencing churn.

## Dataset

* Dataset: Telco Customer Churn Dataset
* Total Records: 7043
* Features: Customer demographics, account information, subscribed services, billing information, and churn status.

## Data Preprocessing

* Removed customerID column
* Converted TotalCharges to numeric format
* Handled missing values in TotalCharges
* Removed duplicate records
* Performed binary encoding and one-hot encoding
* Applied feature scaling using StandardScaler

## Exploratory Data Analysis

Key insights obtained from EDA:

* Customers with month-to-month contracts showed the highest churn rates.
* Electronic check users were more likely to churn.
* Fiber optic customers exhibited significantly higher churn.
* Customers without Online Security, Tech Support, and Online Backup services had higher churn rates.
* Customers with longer tenure were less likely to churn.

## Models Implemented

1. Logistic Regression
2. Decision Tree Classifier
3. Random Forest Classifier

## Model Performance

| Model                       | Accuracy |
| --------------------------- | -------- |
| Logistic Regression         | 82%      |
| Decision Tree (max_depth=5) | 80.6%    |
| Random Forest               | 79.2%    |

### Additional Evaluation Metrics

* ROC-AUC Score: 0.862
* Cross Validation Accuracy: 80.42%
* Precision (Churn): 0.69
* Recall (Churn): 0.60
* F1 Score (Churn): 0.64

## Key Findings

* Contract type is one of the strongest indicators of churn.
* Long-term customers are less likely to churn.
* Additional services such as Online Security and Tech Support improve customer retention.
* Payment method plays a significant role in customer behavior.

## Technologies Used

* Python
* Pandas
* NumPy
* Matplotlib
* Seaborn
* Scikit-Learn
* Jupyter Notebook

## Future Improvements

* Hyperparameter tuning
* XGBoost implementation
* Deployment using Flask or Streamlit
* Real-time churn prediction dashboard
## Project Workflow

1. Data Cleaning
2. Missing Value Handling
3. Duplicate Removal
4. Exploratory Data Analysis (EDA)
5. Feature Engineering
6. Train-Test Split
7. Feature Scaling
8. Logistic Regression
9. Decision Tree
10. Random Forest
11. Model Evaluation
12. Business Insights

## Author

Vishal Prajapati
IIT Bhubaneswar
