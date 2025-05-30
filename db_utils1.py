import pyodbc
import pandas as pd

def connect_db():
    try:
        conn = pyodbc.connect(
            'Driver={SQL Server};'
            'Server=192.168.29.7;'
            'Database=TREASURY_APP;'
            'UID=sa;'
            'PWD=__;'
        )
        return conn
    except Exception as e:
        raise Exception(f"Database connection failed: {e}")

def sync_sheet_to_db(df, table_name="dbo.CURRENCIES", primary_key="id"):
    conn = connect_db()
    cursor = conn.cursor()
    inserted, updated = 0, 0

    for _, row in df.iterrows():
        pk_value = row[primary_key]

        try:
            # Check if row with the same primary key already exists
            cursor.execute(f"SELECT * FROM {table_name} WHERE {primary_key} = ?", pk_value)
            result = cursor.fetchone()

            if result:
                # Convert SQL row result into dictionary for comparison
                db_data = {desc[0]: val for desc, val in zip(cursor.description, result)}
                changes = any(str(row[col]) != str(db_data.get(col)) for col in df.columns if col in db_data)

                if changes:
                    # Prepare update query
                    update_query = f"""
                        UPDATE {table_name} SET
                        {", ".join([f"{col}=?" for col in df.columns if col != primary_key])}
                        WHERE {primary_key}=?
                    """
                    values = [row[col] for col in df.columns if col != primary_key] + [pk_value]
                    cursor.execute(update_query, values)
                    updated += 1
            else:
                # Prepare insert query
                insert_query = f"""
                    INSERT INTO {table_name} ({', '.join(df.columns)})
                    VALUES ({', '.join(['?' for _ in df.columns])})
                """
                cursor.execute(insert_query, [row[col] for col in df.columns])
                inserted += 1

        except Exception as e:
            print(f"❌ Error processing row with {primary_key}={pk_value}: {e}")
            continue

    conn.commit()
    cursor.close()
    conn.close()
    return inserted, updated
