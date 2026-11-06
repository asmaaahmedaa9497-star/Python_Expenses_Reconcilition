import pandas as pd
import os

# ==================================
# 1. تحديد مكان الملفات
# ==================================

folder_path = "/storage/emulated/0/Download"

files = [
 "expenses_january.xlsx",
 "expenses_february.xlsx"
]

all_data = []

# ==================================
# 2. قراءة ملفات المصروفات
# ==================================

for file in files:
 try:
  file_path = os.path.join(folder_path, file)

  df = pd.read_excel(file_path)

  df["source_file"] = file

  all_data.append(df)

 except FileNotFoundError:
  print("File not found:", file)

 except PermissionError:
  print("Permission denied:", file)


if not all_data:
 raise ValueError("No expense files were loaded")

final_df = pd.concat(
 all_data,
 ignore_index=True
)

# ==================================
# 3. قراءة ملف المرجع
# ==================================

reference_file = os.path.join(
 folder_path,
 "transactions_reference.xlsx"
)

reference_df = pd.read_excel(reference_file)

final_df["date"] = pd.to_datetime(
 final_df["date"]
)

reference_df["date"] = pd.to_datetime(
 reference_df["date"]
)

# ==================================
# 4. تجهيز جدول المراجعة
# ==================================

review_df = final_df.copy()

review_df["duplicate_status"] = "No Duplicate"

review_df["reference_status"] = "Unique"

# فحص الصفوف المتطابقة بالكامل

exact_duplicates = review_df[
 review_df.duplicated(keep=False)
]

review_df.loc[
 exact_duplicates.index,
 "duplicate_status"
] = "Exact Duplicate"

# فحص أرقام المرجع المتكررة

duplicate_refs = (
 review_df["reference_number"].notna() &
 review_df["reference_number"].duplicated(keep=False)
)

review_df.loc[
 duplicate_refs,
 "reference_status"
] = "Repeated Reference"

# فحص أرقام المرجع الناقصة

missing_refs = review_df[
 "reference_number"
].isna()

review_df.loc[
 missing_refs,
 "reference_status"
] = "Missing Reference"

# ==================================
# 5. فحص القيم الناقصة
# ==================================

columns_to_check = [
 "date",
 "department",
 "expense_type",
 "amount",
 "payment_method",
 "reference_number"
]

review_df["missing_values"] = ""

for col in columns_to_check:

 missing_mask = review_df[col].isna()

 if missing_mask.any():

  for index in review_df.index[missing_mask]:

   if review_df.at[index, "missing_values"]:

    review_df.at[index, "missing_values"] += ", " + col

   else:

    review_df.at[index, "missing_values"] = col

# ==================================
# 6. دمج المصروفات مع المرجع
# ==================================

review_df = review_df.merge(

 reference_df[
  [
   "reference_number",
   "transaction_id",
   "date",
   "department",
   "expense_type",
   "amount",
   "payment_method"
  ]
 ],

 on="reference_number",

 how="left",

 suffixes=("", "_ref"),

 indicator=True
)

# ==================================
# 7. تنظيف المسافات الزائدة
# ==================================

columns_to_update = [
 "date",
 "department",
 "expense_type",
 "amount",
 "payment_method"
]

for col in columns_to_update:

 review_df[col] = review_df[col].apply(
  lambda x: x.strip() if isinstance(x, str) else x
 )

 review_df[col + "_ref"] = review_df[col + "_ref"].apply(
  lambda x: x.strip() if isinstance(x, str) else x
 )

# ==================================
# 8. تصحيح القيم وتسجيل التغييرات
# ==================================

changes_log = []

for index, row in review_df.iterrows():

 if row["_merge"] != "both":
  continue

 for col in columns_to_update:

  old_value = row[col]

  new_value = row[col + "_ref"]

  if pd.notna(new_value):

   if pd.isna(old_value) or old_value != new_value:

    changes_log.append({

     "reference_number": row["reference_number"],

     "column": col,

     "old_value": old_value,

     "new_value": new_value

    })

    review_df.at[index, col] = new_value


changes_df = pd.DataFrame(
 changes_log,
 columns=[
  "reference_number",
  "column",
  "old_value",
  "new_value"
 ]
)

# ==================================
# 9. تجهيز جدول المصروفات النهائي
# ==================================

corrected_df = review_df[

 [
  "date",
  "department",
  "expense_type",
  "amount",
  "payment_method",
  "reference_number",
  "source_file",
  "duplicate_status",
  "reference_status",
  "missing_values"
 ]

].copy()

# ==================================
# 10. استخراج المعاملات غير المطابقة
# ==================================

unmatched_df = review_df[
 review_df["_merge"] == "left_only"
].copy()

# ==================================
# 11. حفظ التقرير النهائي
# ==================================

output_file = os.path.join(
 folder_path,
 "reconciliation_report.xlsx"
)

with pd.ExcelWriter(
 output_file,
 engine="openpyxl"
) as writer:

 corrected_df.to_excel(
  writer,
  sheet_name="Corrected Expenses",
  index=False
 )

 changes_df.to_excel(
  writer,
  sheet_name="Changes Log",
  index=False
 )

 unmatched_df.to_excel(
  writer,
  sheet_name="Unmatched Expenses",
  index=False
 )

# ==================================
# 12. التأكد من الحفظ
# ==================================

print("Report created successfully!")

print("Changes recorded:", len(changes_df))

print("Unmatched expenses:", len(unmatched_df))

print("File exists:", os.path.exists(output_file))

print("Saved at:", output_file)

