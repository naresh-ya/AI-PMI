#!/usr/bin/env python3
"""
Update the EXISTING dev-ca QA migration inventory workbook IN PLACE.

Rules:
- Never overwrite a cell that already has a value (preserves any Y/N entries).
- Append new rows to existing sheets only when a resource isn't already listed.
- Add brand-new sheets (IAM Policies, KMS, CodeBuild, CodeCommit, Amplify).
Header row = 4, data starts row 5 (consistent across all sheets).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PATH = r"c:\code\Terraform-PMI\dev-ca_QA_migration_inventory.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
EXCLUDE_FILL = PatternFill("solid", fgColor="FCE4D6")
INCLUDE_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)

HEADER_ROW = 4
DATA_START = 5


def fill_for(rec):
    if isinstance(rec, str):
        u = rec.upper()
        if u.startswith("EXCLUDE"):
            return EXCLUDE_FILL
        if u.startswith("MOVE"):
            return INCLUDE_FILL
        if u.startswith("REVIEW") or u.startswith("DECISION"):
            return WARN_FILL
    return None


def existing_keys(ws, key_col=1):
    keys = set()
    for r in range(DATA_START, ws.max_row + 1):
        v = ws.cell(row=r, column=key_col).value
        if v is not None:
            keys.add(str(v).strip())
    return keys


def append_rows(ws, rows, rec_col_idx, key_col=1):
    """Append rows whose key (col1) is not already present. Preserves everything."""
    have = existing_keys(ws, key_col)
    r = ws.max_row + 1
    added = 0
    for rowvals in rows:
        key = str(rowvals[0]).strip()
        if key in have:
            continue
        fill = fill_for(rowvals[rec_col_idx]) if rec_col_idx is not None else None
        for c, val in enumerate(rowvals, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = WRAP
            cell.border = BORDER
            if fill:
                cell.fill = fill
        r += 1
        added += 1
    return added


def new_sheet(wb, title, headers, rows, widths, rec_col_idx):
    if title in wb.sheetnames:
        return wb[title], 0  # already exists, don't recreate
    ws = wb.create_sheet(title=title)
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    n = ws.cell(row=2, column=1,
                value="Set 'Move to QA?' to Y or N per row. "
                      "Green=move, Orange=exclude, Yellow=decision needed.")
    n.font = Font(italic=True, size=9, color="808080")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=HEADER_ROW, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    r = DATA_START
    for rowvals in rows:
        fill = fill_for(rowvals[rec_col_idx]) if rec_col_idx is not None else None
        for c, val in enumerate(rowvals, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.alignment = WRAP
            cell.border = BORDER
            if fill:
                cell.fill = fill
        r += 1
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=DATA_START, column=1)
    ws.sheet_view.showGridLines = False
    return ws, len(rows)


wb = openpyxl.load_workbook(PATH)
report = []

# --- 1) Enrich existing Lambda sheet: add newly-relevant function pmi-master-reviewer ---
# (columns: Function Name, Runtime, IAM Role, Recommendation, Move to QA?, Notes -> rec idx 3)
lam_new = [
    ["pmi-master-reviewer", "python3.10", "pmi-lambda-execution-role", "REVIEW", "",
     "Invoked by pmi-agentic-ai-workflow-dev via pmi-stepfunctions-invoke-policy. Not in current TF map. Confirm."],
]
report.append(("Lambda +rows", append_rows(wb["Lambda"], lam_new, rec_col_idx=3)))

# --- 2) Enrich existing ECR sheet: add agentcore repo if missing ---
ecr_new = [
    ["bedrock-agentcore-calculator_agent", "EXCLUDE", "", "AgentCore demo image - exclude."],
]
report.append(("ECR +rows", append_rows(wb["ECR"], ecr_new, rec_col_idx=1)))

# --- 3) Enrich IAM Roles sheet with roles/notes discovered in revalidation ---
# columns: Role Name, Recommendation, Move to QA?, Notes  -> rec idx 1
iam_new = [
    ["ca-workbench-api-lambda-role", "EXCLUDE", "", "Workbench role - exclude."],
]
report.append(("IAM Roles +rows", append_rows(wb["IAM Roles"], iam_new, rec_col_idx=1)))

# --- 4) NEW SHEET: IAM Policies (fully-resolved role->policy detail) ---
pol_rows = [
    ["pmi-lambda-execution-role", "INLINE", "pmi-lambda-custom-permissions", "bedrock:InvokeModel*/ListFoundationModels; s3 pmi-agentic-ai-documents-*; DDB requests/hitl/knowledge-base; states:*; logs:*", "MOVE", "", "In Terraform (combined). OK."],
    ["pmi-lambda-execution-role", "INLINE", "ccaiaa-prompt-mgmt-read", "bedrock:GetPrompt, ListPrompts (*)", "MOVE", "", "In Terraform. OK."],
    ["pmi-lambda-execution-role", "INLINE", "kbpolicy", "bedrock:Retrieve/StartIngestionJob/... on KB OSJDPJLTNC (pmi-experimental-knowledge-base)", "DECISION", "", "TF uses var.bedrock_kb_id. Live points to experimental KB. Confirm QA KB id."],
    ["pmi-lambda-execution-role", "INLINE", "LambdaDynamoDBCuratedLinksReadAccess", "dynamodb Get/Scan/Query on pmi-agentic-ai-curated-links + index", "MOVE", "", "In Terraform. OK."],
    ["pmi-lambda-execution-role", "INLINE", "EC2networkinterfacepermissions", "ec2 Create/Describe/DeleteNetworkInterface (*)", "MOVE", "", "In Terraform. OK."],
    ["pmi-lambda-execution-role", "INLINE", "ChatContentKmsDecrypt", "kms:Decrypt/DescribeKey on key 0b34a78a-de26-441c-b051-df6f5022afc6", "DECISION", "", "NOT in Terraform. QA needs equivalent KMS key + grant or encrypted-content breaks."],
    ["pmi-lambda-execution-role", "INLINE", "AgentCoreMemory", "bedrock-agentcore ListEvents/CreateEvent/RetrieveMemoryRecords/ListMemoryRecords (*)", "DECISION", "", "NOT in Terraform ('for testing'). Confirm if needed in QA."],
    ["pmi-lambda-execution-role", "INLINE", "ECRPullPolicy", "ecr GetAuthorizationToken/BatchGetImage/... incl repo ca-workbench", "DECISION", "", "NOT in Terraform. Needed only if a Lambda uses container images."],
    ["pmi-lambda-execution-role", "MANAGED", "AmazonSESFullAccess", "Full SES", "DECISION", "", "NOT in Terraform. Confirm if any Lambda sends email via SES."],
    ["pmi-lambda-execution-role", "MANAGED", "AmazonSNSFullAccess", "Full SNS", "DECISION", "", "NOT in Terraform. No SNS topics exist; confirm."],
    ["pmi-lambda-execution-role", "MANAGED", "AmazonS3FullAccess", "Full S3", "MOVE", "", "In Terraform standard set."],
    ["pmi-lambda-execution-role", "MANAGED", "AmazonDynamoDBFullAccess", "Full DynamoDB", "MOVE", "", "In Terraform standard set."],
    ["pmi-lambda-execution-role", "MANAGED", "AWSLambda_FullAccess", "Full Lambda", "MOVE", "", "In Terraform standard set."],
    ["pmi-lambda-execution-role", "MANAGED", "SecretsManagerReadWrite", "Secrets RW", "MOVE", "", "In Terraform standard set."],
    ["pmi-lambda-execution-role", "MANAGED", "AWSLambdaBasicExecutionRole", "CW logs", "MOVE", "", "In Terraform standard set."],
    ["pmi-lambda-execution-role", "MANAGED", "AWSLambdaVPCAccessExecutionRole", "VPC ENI", "MOVE", "", "In Terraform standard set."],
    ["pmi-ingestion-role", "INLINE", "PmiIngestionPolicy", "s3 sandbox bucket; DDB document_metadata + knowledge-base; bedrock:InvokeModel; secrets pmi/sharepoint/*", "MOVE", "", "Matches TF ingestion_policy. OK."],
    ["pmi-ingestion-role", "MANAGED", "AWSLambdaBasicExecutionRole", "CW logs", "MOVE", "", "OK."],
    ["pmi-stepfunctions-execution-role", "INLINE", "pmi-stepfunctions-invoke-policy", "lambda:InvokeFunction on input-validator/master-researcher/master-drafter/master-reviewer; logs; xray", "MOVE", "", "NOT modeled in current TF (no SFN resources yet). Needed for Step Functions."],
    ["pmi-stepfunctions-execution-role", "MANAGED", "AWSLambdaRole", "Invoke Lambda", "MOVE", "", "OK."],
    ["pmi-agentic-ai-create-request-handler-role-dev", "INLINE", "pmi-agentic-ai-create-request-handler-policy-dev", "DDB requests/workflow/timeline/activity(-dev); S3 primary bucket; SFN pmi-agentic-ai-workflow-dev; logs; xray; cw:PutMetricData", "REVIEW", "", "Confirm handler in QA scope."],
    ["pmi-agentic-ai-create-request-handler-role-dev", "MANAGED", "AWSLambdaVPCAccessExecutionRole", "VPC ENI", "REVIEW", "", ""],
    ["pmi-agentic-ai-uploads-presigned-role-dev", "INLINE", "pmi-agentic-ai-uploads-presigned-policy-dev", "s3:PutObject primary bucket uploads/*; logs", "MOVE", "", "Matches TF uploads_presigned_policy."],
    ["pmi-agentic-ai-uploads-presigned-role-dev", "MANAGED", "AWSLambdaBasicExecutionRole", "CW logs", "MOVE", "", "OK."],
    ["pmi-agentic-ai-lambda-role", "MANAGED", "AmazonBedrockFullAccess", "Full Bedrock", "REVIEW", "", "Used by policy-updates Lambda."],
    ["pmi-agentic-ai-lambda-role", "MANAGED", "AmazonDynamoDBFullAccess", "Full DynamoDB", "REVIEW", "", ""],
    ["pmi-agentic-ai-lambda-role", "MANAGED", "AWSLambdaBasicExecutionRole", "CW logs", "REVIEW", "", ""],
    ["amplify-login-lambda-228803b5", "INLINE", "Amplify-backend", "amplifybackend GetToken/DeleteToken; amplify:GetApp", "MOVE", "", "Matches TF amplify_backend."],
]
_, n = new_sheet(wb, "IAM Policies",
                 ["Attached To (Role)", "Type", "Policy Name", "Permissions Summary", "Recommendation", "Move to QA? (Y/N)", "Notes"],
                 pol_rows, [42, 10, 42, 62, 16, 16, 56], rec_col_idx=4)
report.append(("IAM Policies (new sheet)", n))

# --- 5) NEW SHEET: KMS ---
_, n = new_sheet(wb, "KMS", ["Key / Alias", "Recommendation", "Move to QA? (Y/N)", "Notes"], [
    ["key/0b34a78a-de26-441c-b051-df6f5022afc6", "DECISION", "",
     "Referenced by ChatContentKmsDecrypt on pmi-lambda-execution-role. QA needs an equivalent CMK + decrypt grant, or encrypted-content features break. Identify what this key protects."],
], [46, 16, 18, 80], rec_col_idx=1)
report.append(("KMS (new sheet)", n))

# --- 6) NEW SHEET: CodeBuild ---
_, n = new_sheet(wb, "CodeBuild", ["Project", "Recommendation", "Move to QA? (Y/N)", "Notes"], [
    ["ca-workbench-lambda-build", "REVIEW", "", "Workbench build. Move only if workbench app moves to QA."],
    ["ca-workbench-build", "REVIEW", "", "Workbench build. Move only if workbench app moves to QA."],
], [34, 16, 18, 66], rec_col_idx=1)
report.append(("CodeBuild (new sheet)", n))

# --- 7) NEW SHEET: CodeCommit ---
_, n = new_sheet(wb, "CodeCommit", ["Repository", "Recommendation", "Move to QA? (Y/N)", "Notes"], [
    ["CA-UI", "REVIEW", "", "UI source repo. Source control usually not duplicated per-env; confirm."],
], [30, 16, 18, 66], rec_col_idx=1)
report.append(("CodeCommit (new sheet)", n))

# --- 8) NEW SHEET: Amplify ---
_, n = new_sheet(wb, "Amplify", ["App Name", "App ID (dev)", "Recommendation", "Move to QA? (Y/N)", "Notes"], [
    ["PMICAUI", "dtnpl3xf4hrg2", "MOVE", "", "Matches Cognito pool amplify_backend_manager_dtnpl3xf4hrg2. CCAIAA frontend."],
    ["caaiaa_agentic_app", "drp6k4av42dyw", "REVIEW", "", "Agentic app frontend. Confirm which UI is canonical for QA."],
    ["PMI-React-App", "d3n3mnyox8nrt", "REVIEW", "", "React app. Confirm ownership/scope."],
], [24, 18, 16, 18, 62], rec_col_idx=2)
report.append(("Amplify (new sheet)", n))

# --- 9) Add a note into Other & Not Present about SES/SNS/CodePipeline/SSM if not present ---
other_new = [
    ["CloudFront", "NONE", "No distributions."],
    ["SES identities", "NONE", "None - yet pmi-lambda-execution-role has AmazonSESFullAccess. Confirm."],
    ["CodePipeline", "NONE", "None."],
    ["SSM Parameter Store (app params)", "NONE", "No /pmi or /ccaiaa params found."],
    ["EventBridge Scheduler", "NONE", "No schedules (schedules are classic EventBridge rules)."],
]
report.append(("Other & Not Present +rows", append_rows(wb["Other & Not Present"], other_new, rec_col_idx=None)))

wb.save(PATH)
print("Updated in place:", PATH)
for name, cnt in report:
    print(f"  {name}: {cnt}")
print("Final sheets:", wb.sheetnames)
