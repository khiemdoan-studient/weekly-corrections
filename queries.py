"""BigQuery queries for the Weekly Corrections tool."""


def query_alpha_roster(client, project, dataset, table):
    """Query all students from the alpha_roster BQ table (deduped, all statuses).

    Returns a list of dicts keyed by lowercase field names.
    """
    fqn = f"`{project}.{dataset}.{table}`"
    sql = f"""
    SELECT
        fullid          AS student_id,
        campus,
        gradelevel      AS grade,
        alphalevellong  AS level,
        firstname       AS first_name,
        lastname        AS last_name,
        email,
        student_group,
        advisor         AS guide_name,
        advisoremail    AS guide_email,
        externalstudentid AS ext_student_id,
        admissionstatus
    FROM {fqn}
    """
    print(f"  Querying {fqn}...")
    rows = list(client.query(sql).result())
    print(f"  Got {len(rows):,} SIS students from BigQuery")
    return [dict(r) for r in rows]
