from openpyxl import Workbook

from jira_processor import process_daily_jira_workbook


def test_daily_jira_excludes_removed_users(tmp_path):
    source = tmp_path / "daily-jira.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["key", "status", "summary", "developer", "assignee", "reporter"])
    sheet.append(["TEST-1", "处理中", "保留任务", "linyuping", "", ""])
    sheet.append(["TEST-2", "开放", "移除芮顺子", "ruishunzi", "", ""])
    sheet.append(["TEST-3", "已解决", "移除莫钧友", "mojunyou", "", ""])
    workbook.save(source)
    workbook.close()

    result = process_daily_jira_workbook(source)

    summaries = {item["name"]: item for item in result["summaries"]}
    assert "ruishunzi" not in summaries
    assert "mojunyou" not in summaries
    assert summaries["linyuping"]["task_count"] == 1
    assert result["stats"]["included_rows"] == 1
    assert result["stats"]["excluded_rows"] == 2

    removed_records = [record for record in result["records"] if record["summary_person"] in {"ruishunzi", "mojunyou"}]
    assert all(not record["included"] for record in removed_records)
    assert all(record["consistency"] == "不在每日名单" for record in removed_records)
