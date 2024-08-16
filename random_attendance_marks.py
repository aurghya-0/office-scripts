import pandas as pd
import numpy as np
import random as r

# n_students = Number of students
# n_days = Number of days
def generate_random_attendance(n_students, n_days):
  l = [] # final list of attendance
  for _ in (0, n_days):
    a = r.choices(['P','A'], weights=(90,10), k=n_days)
    l.append(a)

    arr = np.array(l)
    arr = np.transpose(arr)
    pd.DataFrame(arr).to_csv('attendance.csv')


def generate_random_marks(n_students, marks_max, marks_min):
    marks = []
    for _ in range(n_students):
        marks = r.choices(range(marks_min, marks_max), k=n_students)

    return marks


if __name__ == "__main__":
    generated_marks = generate_random_marks(30, 24, 21)
    print(generated_marks)
