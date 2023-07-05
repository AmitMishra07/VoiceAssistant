import face_recognition
import cv2
import numpy as np


class MyFaceRecognitionEngine:
    known_face_encodings =None
    known_faces_names=None
    USER_NAME=None

    def __init__(self):
        self.init_values()

    def init_values(self):
        # get palak's encodings
        palak_image=face_recognition.load_image_file("faces/palak.jpg")
        palak_encodings=face_recognition.face_encodings(palak_image)[0]

        # get Amit's encodings
        amit_image=face_recognition.load_image_file("faces/amit.jpg")
        amit_encodings=face_recognition.face_encodings(amit_image)[0]

        ankur_image=face_recognition.load_image_file("faces/ankur.jpg")
        ankur_encodings = face_recognition.face_encodings(ankur_image)[0]

        known_face_encodings=[amit_encodings,palak_encodings,ankur_encodings]
        known_faces_names=["Amit","Palak","Ankur"]
        self.identify_user(known_face_encodings,known_faces_names)


    def identify_user(self, known_face_encodings,known_faces_names):
        vid = cv2.VideoCapture(0)

        while (True):

            # Capture the video frame
            # by   frame
            ret, frame = vid.read()
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            # small_frame=cv2.resize(frame,None,fx=2,fy=2)
            face_locations = face_recognition.face_locations(rgb_frame)
            face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

            for face_encoding in face_encodings:
                matches = face_recognition.compare_faces(known_face_encodings, face_encoding, 0.6)
                face_distance = face_recognition.face_distance(known_face_encodings, face_encoding)
                best_match_index = np.argmin(face_distance)

                if (matches[best_match_index]):
                    name = known_faces_names[best_match_index]
                    self.USER_NAME=name
                    # if name in known_faces_names:
                    font = cv2.FONT_HERSHEY_COMPLEX

                    cv2.putText(frame, "Hi! " + name, (10, 100), font, 1.5, (255, 0, 0), 3, 2)

            break


        # After the loop release the cap object
        vid.release()
        # Destroy all the windows
        cv2.destroyAllWindows()

    def get_user_name(self):
        return self.USER_NAME
