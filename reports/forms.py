# -*- encoding:utf-8 -*-
from django import forms
from .models import Machine

class ReportForm(forms.Form):
    report_choices = [('daily', '日报表'), ('weekly', '周报表'), ('monthly', '月报表')]
    # report_machine_choice_orin = (Machine.objects.all().values_list('machine_id','machine_id'))
    # report_machine_choice_list = []
    # for x in report_machine_choice_orin:
    #     report_machine_choice_list.append(x)
    报表类型 = forms.ChoiceField(choices=report_choices, widget=forms.RadioSelect())
    # report_date = forms.DateField(widget=forms.SelectDateWidget(), label="Select Date")
    日期 = forms.DateField(widget=forms.DateInput(attrs={'class': 'vDateField'}))
    # report_machine = forms.ChoiceField(choices=report_machine_choice_list,widget=forms.Select())
