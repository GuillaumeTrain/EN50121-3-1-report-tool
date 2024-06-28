import sys
import typing
import csv as csv
# import numpy as np
if 'numpy' not in sys.modules:
    import numpy as np
else:
    np = sys.modules['numpy']

class FileReader:
    filepath: str = None
    delimiter: str = None
    reader: csv.reader = None
    fmin = None
    fmax = None
    values_count = None
    points = np.array([[float, float]])

    def __init__(self, filepath=None):
        self.filepath: str = filepath
        self.delimiter: str = None
        self.reader: csv.reader = None
        self.fmin = None
        self.fmax = None
        self.values_count = None
        self.points = []
        #print(self.filepath)

    @property
    def filepath(self):
        return self._filepath

    @filepath.setter
    def filepath(self, value: str):
        self._filepath = value

    @property
    def delimiter(self):
        return self._delimiter

    @delimiter.setter
    def delimiter(self, value):
        self._delimiter = value

    def get_arrayfromcsv(self, trace_number: int):

        if self.filepath:
            if self.filepath != "":
                fmin = None
                fmax = None
                values_count = None
                self.points = np.empty((0, 2), float)
                trace_number = trace_number
                # Lecture du fichier CSV
                with open(self.filepath, 'r') as file:
                    reader = csv.reader(file, delimiter=';')
                    trace_index = 0
                    i = 0
                    for row in reader:
                        # Supprimez les espaces blancs autour des valeurs
                        row = [i.strip() for i in row]

                        # Cherchez fmin, fmax et count of values
                        #print(f"ezefviuh row0 val :{row[0]}")
                        #print(f"ezefviuh trace_index:{trace_index} ")
                        if row[0] == "Start":
                            fmin = float(row[1])
                        elif row[0] == "Stop":
                            fmax = float(row[1])
                        elif row[0] == "Values":
                            trace_index = trace_index + 1
                            if trace_index == trace_number:
                                values_count = int(row[1])
                                #print(f"ezefviuh values_count:{values_count} ")
                                #print(f"ezefviuh trace_index:{trace_index} ")

                        # Enregistrez les points après le "TRACE 1:"
                        elif trace_index == trace_number and len(row) > 1 and row[0].replace('.', '', 1).isdigit() and i < values_count:

                            # Convertez les données en floats et ajoutez-les à la liste des points
                            i = i + 1
                            newpoint = np.array([[float(row[0]), float(row[1])]])
                            #print(f"ezefviuh newpoint:{newpoint} ")
                            # print(newpoint)
                            self.points = np.append(self.points, newpoint, axis=0)
                            # print (self.points)
                    # print (self.points)
                    file.close()
                    i = 0
                    return self.points

#                print(f"fmin = {fmin}, fmax = {fmax}, values_count = {values_count}")
#                print("Points: ", points)
