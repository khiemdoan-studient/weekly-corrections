CREATE TABLE coachbot_feed.khiem_alpha_roster_export
WITH (
    format            = 'PARQUET',
    write_compression = 'SNAPPY',
    external_location = 's3://prod-academics-studient-athena-results/exports/khiem_alpha_roster/'
) AS
WITH src AS (
    SELECT
        fullid, campus, gradelevel, alphalevellong, firstname, lastname,
        email, "group" AS student_group, advisor, advisoremail,
        externalstudentid, admissionstatus
    FROM studient.alpha_student
),
ranked AS (
    SELECT
        s.*,
        ROW_NUMBER() OVER (
            PARTITION BY s.fullid
            ORDER BY
                CASE WHEN s.admissionstatus = 'Enrolled' THEN 0 ELSE 1 END ASC,
                CASE WHEN s.student_group IS NOT NULL AND s.student_group <> '' THEN 0 ELSE 1 END ASC,
                s.fullid ASC
        ) rn
    FROM src s
)
SELECT
    fullid, campus, gradelevel, alphalevellong, firstname, lastname,
    email, student_group, advisor, advisoremail, externalstudentid, admissionstatus
FROM ranked
WHERE rn = 1
