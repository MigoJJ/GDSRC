# print(real_data_dict)

real_data_dict = {'D_name': '남주희0328-20', 'D_gender': 'f', 'D_age': 54}

import pickle
dictionary_data = real_data_dict
a_file = open("data.pkl", "wb")
pickle.dump(dictionary_data, a_file)
a_file.close()

a_file = open("data.pkl", "rb")
output = pickle.load(a_file)
print(output)
a_file.close()
