class ExcelInPython:
    def __init__(self, arguments):
        self._arguments = {i['uid']: i['value'] for i in arguments}
        
    def _flatten_list(self, subject: list) -> list:
        result = []
        for i in subject:
            if type(i) == list:
                result = result + self._flatten_list(i)
            else:
                result.append(i)
        
        return result
        
    @staticmethod
    def _only_numeric_list(flatten_list: list):
        return [i for i in flatten_list if type(i) in [float, int]]

    def _sum(self, flatten_list: list):
        return sum(self._only_numeric_list(flatten_list))

    def _average(self, flatten_list: list):
        return self._sum(flatten_list)/len(self._only_numeric_list(flatten_list))
        
    def _vlookup(self, lookup_value, table_array: list, col_index_num: int, range_lookup: bool = False):
        # TODO add Range Lookup (https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)
        # TODO search optimization needed
        for row in table_array:
            if row[0] == lookup_value:
                return row[col_index_num - 1]

        return 1  # When the vlookup has not found anything similar to lookup_value, it returns 1 in the excel implementation for this function
        
    def _sum_if(self, range_: list, criteria: callable, sum_range: list = None):
        result = 0
        range_, sum_range = self._flatten_list(range_), self._flatten_list(sum_range)
        for i in range(len(range_)):
            if i < len(sum_range) and criteria(range_[i]):
                result += sum_range[i] or 0
                
        return result

    def exec_function_in(self, cell_uid: str):
        return self.__dict__.get(cell_uid, self.__class__.__dict__[cell_uid])(self)

    def _0_1_20(self):
        return self._arguments.get('_0_1_20', 20000)

    def _0_1_21(self):
        return self._arguments.get('_0_1_21', 0)

    def _0_1_16(self):
        return self._arguments.get('_0_1_16', 0.04)

    def _0_1_3(self):
        return self._arguments.get('_0_1_3', 18379)

    def _2_0_0(self):
        return self._arguments.get('_2_0_0', 'код МОП')

    def _2_1_0(self):
        return self._arguments.get('_2_1_0', 'Оценка от РОП')

    def _2_0_1(self):
        return self._arguments.get('_2_0_1', 20156)

    def _2_1_1(self):
        return self._arguments.get('_2_1_1', 0)

    def _2_0_2(self):
        return self._arguments.get('_2_0_2', 22677)

    def _2_1_2(self):
        return self._arguments.get('_2_1_2', 0.25)

    def _2_0_3(self):
        return self._arguments.get('_2_0_3', 1308)

    def _2_1_3(self):
        return self._arguments.get('_2_1_3', None)

    def _2_0_4(self):
        return self._arguments.get('_2_0_4', 20229)

    def _2_1_4(self):
        return self._arguments.get('_2_1_4', 0.25)

    def _2_0_5(self):
        return self._arguments.get('_2_0_5', 51755)

    def _2_1_5(self):
        return self._arguments.get('_2_1_5', 0.25)

    def _2_0_6(self):
        return self._arguments.get('_2_0_6', 20324)

    def _2_1_6(self):
        return self._arguments.get('_2_1_6', 0)

    def _2_0_7(self):
        return self._arguments.get('_2_0_7', 26813)

    def _2_1_7(self):
        return self._arguments.get('_2_1_7', 0.25)

    def _2_0_8(self):
        return self._arguments.get('_2_0_8', 51995)

    def _2_1_8(self):
        return self._arguments.get('_2_1_8', 0.5)

    def _2_0_9(self):
        return self._arguments.get('_2_0_9', 9714)

    def _2_1_9(self):
        return self._arguments.get('_2_1_9', 0)

    def _2_0_10(self):
        return self._arguments.get('_2_0_10', 50480)

    def _2_1_10(self):
        return self._arguments.get('_2_1_10', 0.25)

    def _2_0_11(self):
        return self._arguments.get('_2_0_11', 18379)

    def _2_1_11(self):
        return self._arguments.get('_2_1_11', 0.25)

    def _2_0_12(self):
        return self._arguments.get('_2_0_12', 43567)

    def _2_1_12(self):
        return self._arguments.get('_2_1_12', 0.5)

    def _2_0_13(self):
        return self._arguments.get('_2_0_13', 50295)

    def _2_1_13(self):
        return self._arguments.get('_2_1_13', 0)

    def _0_1_17(self):
        return self._arguments.get('_0_1_17', self._0_1_17_0()/100)

    def _0_1_18(self):
        return self._arguments.get('_0_1_18', 0)

    def _0_1_19(self):
        return self._arguments.get('_0_1_19', self._0_1_16()+self._0_1_17()+self._0_1_18())

    def _1_5_0(self):
        return self._arguments.get('_1_5_0', 'факт ВВ риелторов для начисления з/п')

    def _1_5_1(self):
        return self._arguments.get('_1_5_1', 12000)

    def _1_5_2(self):
        return self._arguments.get('_1_5_2', 120000)

    def _1_5_3(self):
        return self._arguments.get('_1_5_3', 165500)

    def _1_5_4(self):
        return self._arguments.get('_1_5_4', 221000)

    def _1_5_5(self):
        return self._arguments.get('_1_5_5', 128000)

    def _1_5_6(self):
        return self._arguments.get('_1_5_6', 100880)

    def _1_5_7(self):
        return self._arguments.get('_1_5_7', 107500)

    def _1_5_8(self):
        return self._arguments.get('_1_5_8', 110450)

    def _1_5_9(self):
        return self._arguments.get('_1_5_9', 100000)

    def _1_5_10(self):
        return self._arguments.get('_1_5_10', 137750)

    def _1_5_11(self):
        return self._arguments.get('_1_5_11', 111500)

    def _1_5_12(self):
        return self._arguments.get('_1_5_12', 197450)

    def _1_5_13(self):
        return self._arguments.get('_1_5_13', 135563.1)

    def _1_5_14(self):
        return self._arguments.get('_1_5_14', 55518.3)

    def _1_5_15(self):
        return self._arguments.get('_1_5_15', 93000)

    def _1_5_16(self):
        return self._arguments.get('_1_5_16', 49000)

    def _1_5_17(self):
        return self._arguments.get('_1_5_17', 117612)

    def _1_5_18(self):
        return self._arguments.get('_1_5_18', 164750)

    def _1_5_19(self):
        return self._arguments.get('_1_5_19', 128000)

    def _1_5_20(self):
        return self._arguments.get('_1_5_20', 76000)

    def _1_5_21(self):
        return self._arguments.get('_1_5_21', 100000)

    def _1_5_22(self):
        return self._arguments.get('_1_5_22', 113490)

    def _1_5_23(self):
        return self._arguments.get('_1_5_23', 144500)

    def _1_5_24(self):
        return self._arguments.get('_1_5_24', 164000)

    def _1_5_25(self):
        return self._arguments.get('_1_5_25', 100000)

    def _1_5_26(self):
        return self._arguments.get('_1_5_26', 49000)

    def _1_5_27(self):
        return self._arguments.get('_1_5_27', 161750)

    def _1_5_28(self):
        return self._arguments.get('_1_5_28', 49000)

    def _1_5_29(self):
        return self._arguments.get('_1_5_29', 122759)

    def _1_5_30(self):
        return self._arguments.get('_1_5_30', 143000)

    def _1_5_31(self):
        return self._arguments.get('_1_5_31', None)

    def _1_0_0(self):
        return self._arguments.get('_1_0_0', 'Код риэлтора')

    def _1_0_1(self):
        return self._arguments.get('_1_0_1', 58920)

    def _1_0_2(self):
        return self._arguments.get('_1_0_2', 64035)

    def _1_0_3(self):
        return self._arguments.get('_1_0_3', 64035)

    def _1_0_4(self):
        return self._arguments.get('_1_0_4', 33549)

    def _1_0_5(self):
        return self._arguments.get('_1_0_5', 28871)

    def _1_0_6(self):
        return self._arguments.get('_1_0_6', 39988)

    def _1_0_7(self):
        return self._arguments.get('_1_0_7', 28871)

    def _1_0_8(self):
        return self._arguments.get('_1_0_8', 54398)

    def _1_0_9(self):
        return self._arguments.get('_1_0_9', 70498)

    def _1_0_10(self):
        return self._arguments.get('_1_0_10', 54398)

    def _1_0_11(self):
        return self._arguments.get('_1_0_11', 50083)

    def _1_0_12(self):
        return self._arguments.get('_1_0_12', 50083)

    def _1_0_13(self):
        return self._arguments.get('_1_0_13', 64035)

    def _1_0_14(self):
        return self._arguments.get('_1_0_14', 55909)

    def _1_0_15(self):
        return self._arguments.get('_1_0_15', 39988)

    def _1_0_16(self):
        return self._arguments.get('_1_0_16', 25106)

    def _1_0_17(self):
        return self._arguments.get('_1_0_17', 25106)

    def _1_0_18(self):
        return self._arguments.get('_1_0_18', 59091)

    def _1_0_19(self):
        return self._arguments.get('_1_0_19', 55909)

    def _1_0_20(self):
        return self._arguments.get('_1_0_20', 36614)

    def _1_0_21(self):
        return self._arguments.get('_1_0_21', 54398)

    def _1_0_22(self):
        return self._arguments.get('_1_0_22', 64876)

    def _1_0_23(self):
        return self._arguments.get('_1_0_23', 36614)

    def _1_0_24(self):
        return self._arguments.get('_1_0_24', 70498)

    def _1_0_25(self):
        return self._arguments.get('_1_0_25', 52347)

    def _1_0_26(self):
        return self._arguments.get('_1_0_26', 70498)

    def _1_0_27(self):
        return self._arguments.get('_1_0_27', 70498)

    def _1_0_28(self):
        return self._arguments.get('_1_0_28', 52347)

    def _1_0_29(self):
        return self._arguments.get('_1_0_29', 38446)

    def _1_0_30(self):
        return self._arguments.get('_1_0_30', 33549)

    def _1_0_31(self):
        return self._arguments.get('_1_0_31', None)

    def _0_1_11(self):
        return self._arguments.get('_0_1_11', 0)

    def _0_2_0(self):
        return self._arguments.get('_0_2_0', self._0_1_20()+self._0_1_21()+(self._0_1_19())*(self._0_2_0_2()-self._sum_if(self._1_0_None_0(),self._0_2_0_3(), self._1_5_None_1())+((1) if (self._0_1_11()==1) else (0.5))/100*self._sum_if(self._1_0_None_1(),self._0_2_0_4(), self._1_5_None_2())))

    def _2_0_0_0(self):
        return self._arguments.get('_2_0_0_0', [[self._2_0_0(),self._2_1_0()],[self._2_0_1(),self._2_1_1()],[self._2_0_2(),self._2_1_2()],[self._2_0_3(),self._2_1_3()],[self._2_0_4(),self._2_1_4()],[self._2_0_5(),self._2_1_5()],[self._2_0_6(),self._2_1_6()],[self._2_0_7(),self._2_1_7()],[self._2_0_8(),self._2_1_8()],[self._2_0_9(),self._2_1_9()],[self._2_0_10(),self._2_1_10()],[self._2_0_11(),self._2_1_11()],[self._2_0_12(),self._2_1_12()],[self._2_0_13(),self._2_1_13()]])

    def _0_1_17_0(self):
        return self._arguments.get('_0_1_17_0', self._vlookup(self._0_1_3(), self._2_0_0_0(), 2, 0))

    def _1_5_None_0(self):
        return self._arguments.get('_1_5_None_0', [[self._1_5_0()],[self._1_5_1()],[self._1_5_2()],[self._1_5_3()],[self._1_5_4()],[self._1_5_5()],[self._1_5_6()],[self._1_5_7()],[self._1_5_8()],[self._1_5_9()],[self._1_5_10()],[self._1_5_11()],[self._1_5_12()],[self._1_5_13()],[self._1_5_14()],[self._1_5_15()],[self._1_5_16()],[self._1_5_17()],[self._1_5_18()],[self._1_5_19()],[self._1_5_20()],[self._1_5_21()],[self._1_5_22()],[self._1_5_23()],[self._1_5_24()],[self._1_5_25()],[self._1_5_26()],[self._1_5_27()],[self._1_5_28()],[self._1_5_29()],[self._1_5_30()],[self._1_5_31()]])

    def _1_5_None_1(self):
        return self._arguments.get('_1_5_None_1', [[self._1_5_0()],[self._1_5_1()],[self._1_5_2()],[self._1_5_3()],[self._1_5_4()],[self._1_5_5()],[self._1_5_6()],[self._1_5_7()],[self._1_5_8()],[self._1_5_9()],[self._1_5_10()],[self._1_5_11()],[self._1_5_12()],[self._1_5_13()],[self._1_5_14()],[self._1_5_15()],[self._1_5_16()],[self._1_5_17()],[self._1_5_18()],[self._1_5_19()],[self._1_5_20()],[self._1_5_21()],[self._1_5_22()],[self._1_5_23()],[self._1_5_24()],[self._1_5_25()],[self._1_5_26()],[self._1_5_27()],[self._1_5_28()],[self._1_5_29()],[self._1_5_30()],[self._1_5_31()]])

    def _1_5_None_2(self):
        return self._arguments.get('_1_5_None_2', [[self._1_5_0()],[self._1_5_1()],[self._1_5_2()],[self._1_5_3()],[self._1_5_4()],[self._1_5_5()],[self._1_5_6()],[self._1_5_7()],[self._1_5_8()],[self._1_5_9()],[self._1_5_10()],[self._1_5_11()],[self._1_5_12()],[self._1_5_13()],[self._1_5_14()],[self._1_5_15()],[self._1_5_16()],[self._1_5_17()],[self._1_5_18()],[self._1_5_19()],[self._1_5_20()],[self._1_5_21()],[self._1_5_22()],[self._1_5_23()],[self._1_5_24()],[self._1_5_25()],[self._1_5_26()],[self._1_5_27()],[self._1_5_28()],[self._1_5_29()],[self._1_5_30()],[self._1_5_31()]])

    def _0_2_0_0(self):
        return self._arguments.get('_0_2_0_0', self._flatten_list([self._1_5_None_0()]))

    def _0_2_0_1(self):
        return self._arguments.get('_0_2_0_1', self._only_numeric_list(self._0_2_0_0()))

    def _0_2_0_2(self):
        return self._arguments.get('_0_2_0_2', self._sum(self._0_2_0_1()))

    def _0_2_0_3(self):
        return self._arguments.get('_0_2_0_3', lambda x:x==self._0_1_3())

    def _0_2_0_4(self):
        return self._arguments.get('_0_2_0_4', lambda x:x==self._0_1_3())

    def _1_0_None_0(self):
        return self._arguments.get('_1_0_None_0', [[self._1_0_0()],[self._1_0_1()],[self._1_0_2()],[self._1_0_3()],[self._1_0_4()],[self._1_0_5()],[self._1_0_6()],[self._1_0_7()],[self._1_0_8()],[self._1_0_9()],[self._1_0_10()],[self._1_0_11()],[self._1_0_12()],[self._1_0_13()],[self._1_0_14()],[self._1_0_15()],[self._1_0_16()],[self._1_0_17()],[self._1_0_18()],[self._1_0_19()],[self._1_0_20()],[self._1_0_21()],[self._1_0_22()],[self._1_0_23()],[self._1_0_24()],[self._1_0_25()],[self._1_0_26()],[self._1_0_27()],[self._1_0_28()],[self._1_0_29()],[self._1_0_30()],[self._1_0_31()]])

    def _1_0_None_1(self):
        return self._arguments.get('_1_0_None_1', [[self._1_0_0()],[self._1_0_1()],[self._1_0_2()],[self._1_0_3()],[self._1_0_4()],[self._1_0_5()],[self._1_0_6()],[self._1_0_7()],[self._1_0_8()],[self._1_0_9()],[self._1_0_10()],[self._1_0_11()],[self._1_0_12()],[self._1_0_13()],[self._1_0_14()],[self._1_0_15()],[self._1_0_16()],[self._1_0_17()],[self._1_0_18()],[self._1_0_19()],[self._1_0_20()],[self._1_0_21()],[self._1_0_22()],[self._1_0_23()],[self._1_0_24()],[self._1_0_25()],[self._1_0_26()],[self._1_0_27()],[self._1_0_28()],[self._1_0_29()],[self._1_0_30()],[self._1_0_31()]])
