class ReadData:
    """
    data 폴더 내부의 application.csv 파일과, timetable.xlsx 파일을 읽고 dict형태로 반환하는 함수를 제공한다.
    """
    def __init__(self):
        pass

    def read_application(self) -> dict:
        """
        """
        pass

    def read_timetable(self) -> dict:
        """
        """
        pass

class DivideOnOffLineApplicant:
    """
    """

    application: dict
    timetable: dict

    def __init__(self, application:dict, timetable:dict):
        self.application = application
        self.timetable = timetable

    def divide_online(self) -> None:
        """
        """
        pass

    def divide_offline(self) -> None:
        """
        """
        pass

class CreateData:
    """
    data 폴더에 online_data.xlsx 파일과 offline_data.xlsx 파일을 생성하는 함수를 제공한다.
    기존에 동일한 이름의 파일이 존재할 때에는 덮어쓰기를 한다.
    """

    on_line_applicant_dict: dict
    off_line_applicant_dict: dict

    def __init__(self, on_line_applicant_dict:dict, off_line_applicant_dict:dict):
        self.on_line_applicant_dict = on_line_applicant_dict
        self.off_line_applicant_dict = off_line_applicant_dict

    def create_data(self) -> None:
        pass

def main() -> None:
    """
    main 함수.
    1. ReadData 클래스를 통해 application.csv 파일과 timetable.xlsx 파일을 읽어서 dict 형태로 가져온다.
    2. 이후 timetable.xlsx 파일을 기준으로, application.csv에 있는 지원자의 정보를 오프라인/온라인으로 구분한다
    3. CreateData 클래스를 통해 offline_data.xlsx 파일과 online_data.xlsx 파일을 생성한다.
    """


if __name__ == '__main__':
    pass