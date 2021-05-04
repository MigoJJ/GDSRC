import json

student_data = {
    "1.FirstName": "Gildong",
    "2.LastName": "Hong",
    "3.Age": 20,
    "4.University": "Yonsei University",
    "5.Courses": [
        {
            "Major": "Statistics",
            "Classes": ["Probability",
                        "Generalized Linear Model",
                        "Categorical Data Analysis"]
        },
        {
            "Minor": "ComputerScience",
            "Classes": ["Data Structure",
                        "Programming",
                        "Algorithms"]
        }
    ]
}

# st_json = json.dumps(student_data)
# print(st_json)
#
# st_json2 = json.dumps(student_data, indent=4)
# print(st_json2)
#
# st_json3 = json.dumps(student_data, indent=4, sort_keys=True)
# print(st_json3)


# with open("student_file.json", "w") as json_file:
#     json.dump(student_data, json_file)

with open("student_file.json", "r") as st_json:
    st_python = json.load(st_json)
    print(st_python)



# --- 출처: https://rfriend.tistory.com/474
# --- [R, Python 분석과 프로그래밍의 친구 (by R Friend)]