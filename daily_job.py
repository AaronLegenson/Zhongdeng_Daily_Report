# -*- coding: utf-8 -*-

from tools import *
from mail import Mail
from team_email_book import *
import sys
import argparse
import socket


def create_mail_receivers(legal_list, mod, weekday, content_type, team_mail=None):
    """
    根据各种参数生成对应的发送目标
    :param legal_list: 设定发送日，默认为每周一三五
    :param mod: "dev", "prd", "test", 其中"dev"为08:00启动，"prd"为11:00启动，"test"为任意时间手动启动，正在debug的同事可以将tester_email填成自己
    :param weekday: 1-7，其中分一三五和二四六的不同逻辑
    :param content_type: "team"表示精简表版本, "all"表示全表版本
    :param team_mail: 团队邮箱
    :return: 返回to_receivers, cc_receivers, bcc_receivers三个邮箱字符串列表，返回subject_tail表示subject需要拼接的字符串
    """
    # print(legal_list, mod, weekday, content_type, team_mail)
    to_receivers, cc_receivers, bcc_receivers = [], [], []
    subject_tail = ""
    tester_email = "jinyanxu@fusionfintrade.com"
    # print("\nReceivers:", to_receivers, cc_receivers, bcc_receivers)
    if mod == "test":
        to_receivers = [tester_email]
        cc_receivers = [tester_email]
        bcc_receivers = [tester_email]
        subject_tail = " [开发测试]"
        return to_receivers, cc_receivers, bcc_receivers, subject_tail
    if content_type == "team":  # 精简表
        if mod == "prd" and weekday in legal_list:
            to_receivers = team_mail
            cc_receivers = ["wenghuiying@fusionfintrade.com"]
            bcc_receivers = ["wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
            # subject_tail = ""
    elif content_type == "all":  # 全表
        if mod == "dev":  # 8点档
            if weekday in legal_list:  # 每周一三五
                to_receivers = []
                cc_receivers = []
                bcc_receivers = ["wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
                subject_tail = " [推送日-测试时间-全表版本]"
            else:  # 每周二四六日
                to_receivers = []
                cc_receivers = []
                bcc_receivers = ["wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
                subject_tail = " [非推送日-测试时间-全表版本]"
        elif mod == "prd":  # 11点档
            if weekday in legal_list:  # 每周一三五
                to_receivers = ["shenzhihua@fusionfintrade.com", "renchen@fusionfintrade.com"]
                cc_receivers = []
                bcc_receivers = ["wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
                subject_tail = " [推送日-推送时间-全表版本]"
            else:  # 每周二四六日
                to_receivers = ["shenzhihua@fusionfintrade.com", "renchen@fusionfintrade.com"]
                cc_receivers = []
                bcc_receivers = ["wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
                subject_tail = " [非推送日-推送时间-全表版本]"
    return to_receivers, cc_receivers, bcc_receivers, subject_tail


def daily_job(mod="test"):
    execute_path = sys.path[0]
    log = Log()
    legal_list = [1, 3, 5]

    # Step 0: Basic Information
    log.print("[Step 0] Basic Information: Mod: {0}".format(mod))
    log.print("[Step 0] Basic Information: Weekdays sending to teams: {0}".format(str(legal_list)))
    weekday = get_week_day(now_time_string())
    log.print("[Step 0] Basic Information: Weekday: {0}".format(weekday))
    try:
        server_ip = socket.gethostbyname(socket.gethostname())
    except Exception as e:
        log.print("[Step 0] Error in getting ip:", e)
        server_ip = "unknown"
    log.print("[Step 0] Basic Information: Server IP: {0}".format(server_ip))
    log.print("[Step 0] Basic Information: Done")

    # Step 1: Get Normal Info
    log.print("[Step 1] Get Normal Info: From MySQL at 172.16.24.74")
    mc = MySQLConnection()
    mc.connect()
    df_business = mc.get_business()
    log.print("[Step 1] Get Normal Info: {0} business".format(len(df_business)))
    df_user_outer, df_user_inner = mc.get_user()
    log.print("[Step 1] Get Normal Info: {0} users all / {1} outer users / {2} inner users".format(
        len(df_user_outer) + len(df_user_inner), len(df_user_outer), len(df_user_inner)))
    df_log_all = mc.get_log()
    log.print("[Step 1] Get Normal Info: {0} user query logs".format(len(df_log_all)))
    df_enterprise_register_info = mc.get_enterprise_register_info()
    mc.close()
    log.print("[Step 1] Get Normal Info: Done")

    # Step 2: Build Outer Details
    log.print("[Step 2] Build Outer Details: ", "")
    df_outer = df_log_all.merge(df_user_outer, on=STRING_COL_USER_ID)
    df_outer.sort_values(by=[STRING_COL_QUERY_TIME], ascending=[False], inplace=True)
    df_outer.reset_index(drop=True, inplace=True)
    df_outer.index += 1
    df_outer = df_outer[[
        STRING_COL_ORG_NAME,
        STRING_COL_QUERY_TIME,
        STRING_COL_KEYWORD_COUNT,
        STRING_COL_BUSINESS_ID,
        STRING_COL_BUSINESS_NAME,
        STRING_COL_USER_ID,
        STRING_COL_USER_NAME,
        STRING_COL_TEAM,
        STRING_COL_TEAM_CONTACT
    ]]
    df_newly_added_outer, newly_added_previous_string, newly_added_today_string = build_newly_added(df_outer, legal_list)
    log.print("{0} outer queries / {1} outer newly added".format(len(df_outer), len(df_newly_added_outer)))
    log.print("[Step 2] Build Outer Details: Done")

    # Step 3: Build Inner Details
    log.print("[Step 3] Build Inner Details: ", "")
    df_inner = df_log_all.merge(df_user_inner, on=STRING_COL_USER_ID)
    df_inner.sort_values(by=[STRING_COL_QUERY_TIME], ascending=[False], inplace=True)
    df_inner.reset_index(drop=True, inplace=True)
    df_inner.index += 1
    df_inner = df_inner[[
        STRING_COL_ORG_NAME,
        STRING_COL_QUERY_TIME,
        STRING_COL_KEYWORD_COUNT,
        STRING_COL_BUSINESS_ID,
        STRING_COL_BUSINESS_NAME,
        STRING_COL_USER_ID,
        STRING_COL_USER_NAME,
        STRING_COL_TEAM,
        STRING_COL_TEAM_CONTACT
    ]]
    df_newly_added_inner, newly_added_previous_string, newly_added_today_string = build_newly_added(df_inner,
                                                                                                    legal_list)
    log.print("{0} inner queries / {1} inner newly added".format(len(df_inner), len(df_newly_added_inner)))
    log.print("[Step 3] Build Inner Details: Done")

    # Step 4: Build Report
    save_path = "{0}{1}saves{2}{3}-{4}_{5}.xlsx".format(
        execute_path, os.sep, os.sep, STRING_MAIL_TITLE, STRING_CN_ENTIRE_COMPANY, stamp_to_string(time.time(), "%Y%m%d_%H%M%S"))
    writer = pd.ExcelWriter(save_path)
    content_list = []
    # 准备工作
    df_activity, dic_activity, dic_activity_all = build_activity(df_outer)
    # org_name_list = list(set(df_outer[STRING_COL_ORG_NAME]))
    # org_tab_info_dic = get_tab_detail_new(org_name_list)
    business_names = list(set(df_outer[STRING_COL_BUSINESS_NAME]))
    df_tab_dict = dict()
    for one_business in business_names:
        df_tab = build_tab(df_outer, one_business, df_enterprise_register_info)
        df_tab_dict[one_business] = df_tab
    # 4.0 说明
    df_instructions = build_instructions()
    df_instructions.to_excel(writer, sheet_name=STRING_SHEET_INSTRUCTIONS, index=False)
    log.print("[Step 4] Build Report: 4.0 说明 Done")
    # 4.1 开通机构信息
    df_business_full = build_df_business_full(df_business, df_tab_dict, dic_activity, dic_activity_all)
    df_business_full.to_excel(writer, sheet_name=STRING_SHEET_BUSINESS, index=False)
    content_list.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_BUSINESS, df_business_full.to_html()])
    log.print("[Step 4] Build Report: 4.1 开通机构信息 Done")
    # 4.2 近十天机构查询活跃度排名
    df_activity.to_excel(writer, sheet_name=STRING_SHEET_ACTIVITY, index=False)
    content_list.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_ACTIVITY, df_activity.to_html()])
    log.print("[Step 4] Build Report: 4.2 近十天机构查询活跃度排名 Done")
    # 4.3.1 开通机构用户数据 外部
    df_user_outer.to_excel(writer, sheet_name=STRING_SHEET_USER + "(机构用户)", index=False)
    content_list.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_USER + " [机构用户]", df_user_outer.to_html()])
    log.print("[Step 4] Build Report: 4.3.1 开通机构用户数据 外部 Done")
    # 4.3.2 开通机构用户数据 内部
    df_user_inner.to_excel(writer, sheet_name=STRING_SHEET_USER + "(内部员工)", index=False)
    log.print("[Step 4] Build Report: 4.3.2 开通机构用户数据 内部 Done")
    # 4.4.1 本期新增查询日志 外部
    df_newly_added_outer.to_excel(writer, sheet_name=STRING_SHEET_NEWLY_ADDED_EMPTY + "(机构用户)", index=False)
    if len(df_newly_added_outer) > 0:
        content_list.append(["highlight", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_NEWLY_ADDED.format(newly_added_previous_string, newly_added_today_string) + " [机构用户]", df_newly_added_outer.to_html(float_format="{0:.0f}".format)])
    else:
        content_list.append(["none_red", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_NEWLY_ADDED.format(newly_added_previous_string, newly_added_today_string) + " [机构用户]", "[本期机构用户日志暂无新增]"])
    log.print("[Step 4] Build Report: 4.4.1 本期新增查询日志 外部 Done")
    # 4.4.2 本期新增查询日志 内部
    df_newly_added_inner.to_excel(writer, sheet_name=STRING_SHEET_NEWLY_ADDED_EMPTY + "(内部员工)", index=False)
    log.print("[Step 4] Build Report: 4.4.2 本期新增查询日志 内部 Done")
    # 4.5.1 全量查询日志 外部
    df_outer.to_excel(writer, sheet_name=STRING_SHEET_DETAILS + "(机构用户)", index=False)
    # content_list.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_DETAILS + " [机构用户]", df_outer.to_html()])
    log.print("[Step 4] Build Report: 4.5.1 全量查询日志 外部 Done")
    # 4.5.2 全量查询日志 内部
    df_inner.to_excel(writer, sheet_name=STRING_SHEET_DETAILS + "(内部员工)", index=False)
    # content_list.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_DETAILS + " [内部员工]", df_inner.to_html()])
    log.print("[Step 4] Build Report: 4.5.2 全量查询日志 内部 Done")
    # 4.6 各家机构 仅外部
    for one_business in business_names:
        one_business_clear = remove_brackets(one_business)
        df_tab = df_tab_dict[one_business]
        df_tab.to_excel(writer, sheet_name=one_business_clear, index=False)
        content_list.append(["normal", one_business, df_tab.to_html()])
    log.print("[Step 4] Build Report: 4.6 各家机构 仅外部 Done")
    log.print("[Step 4] Build Report: Saving ", "")
    writer.save()
    log.print("Done")
    log.print("[Step 4] Build Report: Setting Styles ", "")
    set_excel_style(save_path, business_names)
    log.print("Done")
    log.print("[Step 4] Build Report: Save to {0}".format(save_path.replace(execute_path, "{Path}")))
    log.print("[Step 4] Build Report: Done")

    # Step 5: SHA256 Hash
    log.print("[Step 5] SHA256 Hash: ", "")
    file_md5 = sha256(save_path)
    log.print(file_md5)
    log.print("[Step 5] SHA256 Hash: Done")

    # Step 6: Mail Team Reports
    # teams = list(set(df_business[STRING_COL_TEAM]))
    # teams.sort(key=lambda x: x)

    teams = TEAM_LIST
    static_time_string = now_time_string(STRING_FORMAT_NO_HYPHEN)
    for one_team in teams:
        log.print("[Step 6] Mail Team Reports: {0}".format(one_team))
        save_path_team = save_path.replace(STRING_CN_ENTIRE_COMPANY, one_team)
        writer_team = pd.ExcelWriter(save_path_team)
        content_list_team = []
        # # 6.0 说明
        # df_instructions.to_excel(writer_team, sheet_name=STRING_SHEET_INSTRUCTIONS, index=False)
        # 6.1 开通机构信息
        df_business = df_business[[
            STRING_COL_BUSINESS_NAME,
            STRING_COL_ACCOUNT_TYPE,
            STRING_COL_ACCOUNT_COUNT,
            # STRING_COL_WHETHER_FP,
            STRING_COL_RELEASE_DATE,
            STRING_COL_TEAM,
            STRING_COL_TEAM_CONTACT
        ]]
        df_business.to_excel(writer_team, sheet_name=STRING_SHEET_BUSINESS, index=False)
        content_list_team.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_BUSINESS, df_business.to_html()])
        # 6.1.x 近十天机构查询活跃度排名
        df_activity = df_activity[[STRING_COL_RANK, STRING_COL_BUSINESS_NAME, STRING_COL_TEAM]]
        df_activity.to_excel(writer_team, sheet_name=STRING_SHEET_ACTIVITY, index=False)
        content_list_team.append(["normal", STRING_CN_ENTIRE_COMPANY + " " + STRING_SHEET_ACTIVITY, df_activity.to_html()])
        # # 6.2.1 开通机构用户数据 外部
        # df_user_outer_team = df_user_outer[df_user_outer[STRING_COL_TEAM] == one_team]
        # df_user_outer_team.reset_index(drop=True, inplace=True)
        # df_user_outer_team.index += 1
        # df_user_outer_team.to_excel(writer_team, sheet_name=STRING_SHEET_USER + "(机构用户)", index=False)
        # if len(df_user_outer_team) > 0:
        #     content_list_team.append(["normal", one_team + " " + STRING_SHEET_USER + " [机构用户]", df_user_outer_team.to_html()])
        # else:
        #     content_list_team.append(["none_blue", one_team + " " + STRING_SHEET_USER + " [机构用户]", "[暂无已开通的机构用户]"])
        # # 6.2.2 开通机构用户数据 内部
        # df_user_inner_team = df_user_inner[df_user_inner[STRING_COL_TEAM] == one_team]
        # df_user_inner_team.reset_index(drop=True, inplace=True)
        # df_user_inner_team.index += 1
        # df_user_inner_team.to_excel(writer_team, sheet_name=STRING_SHEET_USER + "(内部员工)", index=False)
        # if len(df_user_inner_team) > 0:
        #     content_list_team.append(["normal", one_team + " " + STRING_SHEET_USER + " [内部员工]", df_user_inner_team.to_html()])
        # else:
        #     content_list_team.append(["none_blue", one_team + " " + STRING_SHEET_USER + " [内部员工]", "[暂无已开通的内部员工]"])
        # # 6.3.1 本期新增查询日志 外部
        # df_newly_added_outer_team = df_newly_added_outer[df_newly_added_outer[STRING_COL_TEAM] == one_team]
        # df_newly_added_outer_team.reset_index(drop=True, inplace=True)
        # df_newly_added_outer_team.index += 1
        # df_newly_added_outer_team.to_excel(writer_team, sheet_name=STRING_SHEET_NEWLY_ADDED_EMPTY + "(机构用户)", index=False)
        # if len(df_newly_added_outer_team) > 0:
        #     content_list_team.append(["highlight", one_team + " " + STRING_SHEET_NEWLY_ADDED.format(
        #         newly_added_previous_string,
        #         newly_added_today_string) + " [机构用户]",
        #         df_newly_added_outer_team.to_html(float_format="{0:.0f}".format)])
        # else:
        #     content_list_team.append(
        #         ["none_red", one_team + " " + STRING_SHEET_NEWLY_ADDED.format(newly_added_previous_string, newly_added_today_string) + " [机构用户]", "[本期机构用户日志暂无新增]"])
        # # 6.3.1 本期新增查询日志 内部
        # df_newly_added_inner_team = df_newly_added_inner[df_newly_added_inner[STRING_COL_TEAM] == one_team]
        # df_newly_added_inner_team.reset_index(drop=True, inplace=True)
        # df_newly_added_inner_team.index += 1
        # df_newly_added_inner_team.to_excel(writer_team, sheet_name=STRING_SHEET_NEWLY_ADDED_EMPTY + "(内部员工)", index=False)
        # # 6.4.1 全量查询日志 外部
        # df_outer_team = df_outer[df_outer[STRING_COL_TEAM] == one_team]
        # df_outer_team.reset_index(drop=True, inplace=True)
        # df_outer_team.index += 1
        # df_outer_team.to_excel(writer_team, sheet_name=STRING_SHEET_DETAILS + "(机构用户)", index=False)
        # if len(df_outer_team) > 0:
        #     content_list_team.append(["normal", one_team + " " + STRING_SHEET_DETAILS + " [机构用户]", df_outer_team.to_html()])
        # else:
        #     content_list_team.append(["none_blue", one_team + " " + STRING_SHEET_DETAILS + " [机构用户]", "[机构用户日志暂无]"])
        # # 6.4.1 全量查询日志 内部
        # df_inner_team = df_inner[df_inner[STRING_COL_TEAM] == one_team]
        # df_inner_team.reset_index(drop=True, inplace=True)
        # df_inner_team.index += 1
        # df_inner_team.to_excel(writer_team, sheet_name=STRING_SHEET_DETAILS + "(内部员工)", index=False)
        # # 6.5 各家机构 仅外部
        # business_names_team = list(set(df_outer_team[STRING_COL_BUSINESS_NAME]))
        # for one_business in business_names_team:
        #     # print(one_team, business_names_team)
        #     one_business_clear = remove_brackets(one_business)
        #     df_tab = df_tab_dict.get(one_business)
        #     df_tab.to_excel(writer_team, sheet_name=one_business_clear, index=False)
        #     content_list_team.append(["normal", one_business, df_tab.to_html()])
        writer_team.save()
        set_excel_style(save_path_team, [])
        mail_team = Mail()
        subject_team = "{0}-{1} [{2}]".format(STRING_MAIL_TITLE, one_team, static_time_string)
        # if mod == "test":
        #     to_receivers_team = ["xuenze@fusionfintrade.com"]
        #     cc_receivers_team = ["xuenze@fusionfintrade.com"]
        #     bcc_receivers_team = ["xuenze@fusionfintrade.com"]
        #     subject_team += " [开发测试]"
        # elif get_week_day(now_time_string()) in legal_list:
        #     if mod == "prd":  # 一三五的11点档
        #         to_receivers_team = TEAM_EMAIL_DICTIONARY.get(one_team)
        #         cc_receivers_team = ["wenghuiying@fusionfintrade.com"]
        #         bcc_receivers_team = ["xuenze@fusionfintrade.com", "gusongtao@fusionfintrade.com"]
        #     else:  # 一三五的8点档
        #         to_receivers_team = ["wenghuiying@fusionfintrade.com"]
        #         cc_receivers_team = ["xuenze@fusionfintrade.com"]
        #         bcc_receivers_team = ["xuenze@fusionfintrade.com", "gusongtao@fusionfintrade.com"]
        #         subject_team += " [发送日8点预测试-对公版本]"
        # else:  # 非一三五
        #     if mod == "prd":  # 二四六日的11点档
        #         to_receivers_team = ["xuenze@fusionfintrade.com"]
        #         cc_receivers_team = ["xuenze@fusionfintrade.com"]
        #         bcc_receivers_team = ["xuenze@fusionfintrade.com", "gusongtao@fusionfintrade.com"]
        #         subject_team += " [非一三五自测-对公版本]"
        #     else:  # 二四六日的8点档
        #         to_receivers_team = ["wenghuiying@fusionfintrade.com"]
        #         cc_receivers_team = ["xuenze@fusionfintrade.com"]
        #         bcc_receivers_team = ["xuenze@fusionfintrade.com"]
        #         subject_team += " [非一三五自测-对公版本]"
        to_receivers_team, cc_receivers_team, bcc_receivers_team, subject_tail_team = create_mail_receivers(legal_list, mod, weekday, "team", TEAM_EMAIL_DICTIONARY.get(one_team))
        subject_team += subject_tail_team
        mail_team.set_receivers(to_receivers_team, cc_receivers_team, bcc_receivers_team)
        content_html_team = unite_html(content_list_team, one_team)
        mail_team.send(content_html_team, [save_path_team], subject_team, STRING_HTML)
    log.print("[Step 6] Mail Team Reports: Done")

    # Step 7: Mail Main Report
    mail = Mail()
    log.print("[Step 7] Mail Main Report: Weekday is {0} (1-7)".format(get_week_day(now_time_string())))
    subject = "{0}-{1} [{2}]".format(STRING_MAIL_TITLE, STRING_CN_ENTIRE_COMPANY, static_time_string)
    # if mod == "test":
    #     to_receivers = ["xuenze@fusionfintrade.com"]
    #     cc_receivers = ["xuenze@fusionfintrade.com"]
    #     bcc_receivers = ["xuenze@fusionfintrade.com"]
    #     subject += " [开发测试]"
    #     mail.set_receivers(to_receivers, cc_receivers, bcc_receivers)
    #     log.print("[Step 7] Mail Main Report: Send to {0}".format(str(to_receivers)))
    #     log.print("[Step 7] Mail Main Report: Cc to {0}".format(str(cc_receivers)))
    # elif mod == "prd":  # 11点档：11点不做完整版本，直接返回
    #     return True
    # else:  # 8点档
    #     if get_week_day(now_time_string()) in legal_list:  # 一三五8点
    #         to_receivers = []
    #         cc_receivers = []
    #         bcc_receivers = ["xuenze@fusionfintrade.com", "wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
    #     else:  # 非一三五8点
    #         to_receivers = []
    #         cc_receivers = []
    #         bcc_receivers = ["xuenze@fusionfintrade.com", "wenghuiying@fusionfintrade.com", "gusongtao@fusionfintrade.com", "jinyanxu@fusionfintrade.com"]
    #         subject += " [非一三五自测-08:00-全表版本]"
    #     mail.set_receivers(to_receivers, cc_receivers, bcc_receivers)
    to_receivers, cc_receivers, bcc_receivers, subject_tail = create_mail_receivers(legal_list, mod, weekday, "all")
    subject += subject_tail
    mail.set_receivers(to_receivers, cc_receivers, bcc_receivers)
    log.print("[Step 7] Mail Main Report: Send to {0}".format(str(to_receivers)))
    log.print("[Step 7] Mail Main Report: Cc to {0}".format(str(cc_receivers)))
    content_list.append(["log", "Logs", log.log])
    content_html = unite_html(content_list, STRING_CN_ENTIRE_COMPANY)
    mail.send(content_html, [save_path], subject, STRING_HTML)
    log.print("[Step 7] Mail Main Report: Done")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--mod", default="test", help="operation mod: 'test', 'prd' or 'dev'. default value is 'test'")
    opt = parser.parse_args()
    if opt.mod == "test":
        res = daily_job("test")
    else:
        res = daily_job(opt.mod)

