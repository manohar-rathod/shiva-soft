import requests
import Utilities.excelutil as excel_util
import Utilities.generic as generic


def create_dictionary(header_in_string):
    header = {}
    header_value = header_in_string.replace("#", ',').replace(":", ',').split(",")
    key = []
    value = []
    for i in range(len(header_value)):
        if (i + 1) % 2 == 0:
            value.append(header_value[i])
        else:
            key.append(header_value[i])
    count = 0
    for iterative_data in key:
        header.update({iterative_data: value[count]})
        count += 1
    return header


def test_execute_api(excel_path, web_server_ip):
    data_frame = excel_util.read_excel(excel_path)
    sheet_name = data_frame[1]

    count = 0

    for iterative_data in data_frame[0][sheet_name]:
        api_name = data_frame[0][sheet_name]['ApiName'][count]
        http_method = data_frame[0][sheet_name]['HttpMethod'][count]
        url = web_server_ip + data_frame[0][sheet_name]['RequestURL'][count]
        body = data_frame[0][sheet_name]['BodyParam'][count]
        header = create_dictionary(data_frame[0][sheet_name]['Header'][count])
        status = data_frame[0][sheet_name]['Status'][count]
        response = data_frame[0][sheet_name]['API Response'][count]
        api_response = ''
        dependent_method = data_frame[0][sheet_name]['DependsOnMethod'][count]
        dependent_variable = data_frame[0][sheet_name]['DependsOnGlobalVal'][count]
        if 'post' in http_method.lower():
            print("url : ", url, "data : ", body, " Headers : ", str(header))
            try:

                if dependent_method.lower() != 'none':
                    body = generic.return_body_with_replacing_dependent_variable(data_frame=data_frame[0],
                                                                                 sheet_name=sheet_name,
                                                                                 dependent_method=dependent_method,
                                                                                 dependnet_variable_name=dependent_variable,
                                                                                 body_parameter=body,
                                                                                 column_name='ApiName',
                                                                                 column_name_with_api_response='API Response')
                    print("Global value replaced and body paremeter is : ", body)

                api_response = requests.post(url=url, data=body, headers=header)
                print("Api Response : ", api_response.text, "API status code", api_response.status_code)
                if api_response.status_code == 200:
                    excel_util.update_cell_value(data_frame=data_frame[0], sheet_name=sheet_name,
                                                 count=count, api_response=api_response.text, status='Pass',
                                                 status_code=api_response.status_code)
                else:
                    excel_util.update_cell_value(data_frame=data_frame[0], sheet_name=sheet_name,
                                                 count=count, api_response=api_response.text, status='Fail',
                                                 status_code=api_response.status_code)
            except Exception as e:
                excel_util.update_cell_value(data_frame=data_frame[0], sheet_name=sheet_name,
                                             count=count, api_response=e, status='ERROR',
                                             status_code='500')

        count += 1
        if len(data_frame[0][sheet_name]['ApiName'].index.tolist()) == count:
            break

    excel_util.write_excel(data_frame[0][sheet_name], sheet_name,
                           'D:\\Software\\Pycharm\\PycharmProjects\\APIAutomation\\output\\')


# create_dictionary("Content-Type:application/json#TenantId:gbl_TenantId#UserId:83485#Token:abcd")
test_execute_api("D:\\Software\\Pycharm\\PycharmProjects\\APIAutomation\\TestData\\", 'http://10.1.172.123:1005/')
