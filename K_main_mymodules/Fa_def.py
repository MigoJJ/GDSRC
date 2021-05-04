import pickle

def pk():
    real_data_dict = {'R_name': jj, 'R_date': C_date, 'R_gender': gender}
    dictionary_data = real_data_dict
    with open("data.pkl", "wb") as a_file:
        pickle.dump(dictionary_data, a_file)
    a_file.close()