import calendar
from datetime import date, datetime


class CalendarUtil:
    def __init__(self, year=date.today().year, month=date.today().month):
        self.year = year
        self.month = month
        self.weekdays = [
            'Segunda-Feira',
            'Terça-Feira',
            'Quarta-Feira',
            'Quinta-Feira',
            'Sexta-Feira',
            'Sábado',
            'Domingo'
        ]
        self.months = [
            'Janeiro',
            'Fevereiro',
            'Março',
            'Abril',
            'Maio',
            'Junho',
            'Julho',
            'Agosto',
            'Setembro',
            'Outubro',
            'Novembro',
            'Dezembro,'
        ]


    def get_current_month_range(self):
        _, day_count = calendar.monthrange(self.year, self.month)

        return [i+1 for i in range(day_count)]
    
    def get_weekday(self, day):
        date_obj = date(self.year, self.month, day)
        return (date_obj.weekday(), self.weekdays[date_obj.weekday()])


if __name__ == '__main__':
    print(CalendarUtil().get_current_month_range())