import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { pool } from "../lib/db.js";

const listQuerySchema = z.object({
  query: z.string().trim().max(120).optional(),
  department: z.string().trim().max(120).optional(),
  status: z
    .enum(["active", "onboarding", "inactive", "terminated"])
    .optional(),
  limit: z.coerce.number().int().min(1).max(100).default(25),
  offset: z.coerce.number().int().min(0).default(0),
});

const employeeParamsSchema = z.object({
  id: z.string().trim().min(1).max(80),
});

type EmployeeRow = {
  id: string;
  employee_code: string;
  full_name: string;
  status: "active" | "onboarding" | "inactive" | "terminated";
  birth_day: number | null;
  birth_month: number | null;
  birth_year: number | null;
  hired_on: string | null;
  terminated_on: string | null;
  notes: string | null;
  position_id: string | null;
  position_code: string | null;
  position_title: string | null;
  department_id: string | null;
  department_code: string | null;
  department_name: string | null;
  department_location: string | null;
  contacts: Array<{
    type: string;
    value: string;
    isPrimary: boolean;
  }>;
  total_count: number;
};

const employeeSelectSql = `
  SELECT
    employee.id::text,
    employee.employee_code,
    employee.full_name,
    employee.status,
    employee.birth_day,
    employee.birth_month,
    employee.birth_year,
    employee.hired_on::text,
    employee.terminated_on::text,
    employee.notes,
    position.id::text AS position_id,
    position.code AS position_code,
    position.title AS position_title,
    department.id::text AS department_id,
    department.code AS department_code,
    department.name AS department_name,
    department.location AS department_location,
    COALESCE(contact_data.contacts, '[]'::jsonb) AS contacts,
    count(*) OVER()::int AS total_count
  FROM employees AS employee
  LEFT JOIN employment_assignments AS assignment
    ON assignment.employee_id = employee.id
   AND assignment.ended_on IS NULL
   AND assignment.is_primary = TRUE
  LEFT JOIN positions AS position
    ON position.id = assignment.position_id
  LEFT JOIN departments AS department
    ON department.id = position.department_id
  LEFT JOIN LATERAL (
    SELECT jsonb_agg(
      jsonb_build_object(
        'type', contact.contact_type,
        'value', contact.contact_value,
        'isPrimary', contact.is_primary
      )
      ORDER BY contact.is_primary DESC, contact.contact_type, contact.contact_value
    ) AS contacts
    FROM employee_contacts AS contact
    WHERE contact.employee_id = employee.id
      AND contact.is_active = TRUE
  ) AS contact_data ON TRUE
`;

function mapEmployee(row: EmployeeRow) {
  return {
    id: row.id,
    employeeCode: row.employee_code,
    fullName: row.full_name,
    status: row.status,
    birthday:
      row.birth_day !== null && row.birth_month !== null
        ? {
            day: row.birth_day,
            month: row.birth_month,
            year: row.birth_year,
          }
        : null,
    hiredOn: row.hired_on,
    terminatedOn: row.terminated_on,
    notes: row.notes,
    position: row.position_id
      ? {
          id: row.position_id,
          code: row.position_code,
          title: row.position_title,
        }
      : null,
    department: row.department_id
      ? {
          id: row.department_id,
          code: row.department_code,
          name: row.department_name,
          location: row.department_location,
        }
      : null,
    contacts: row.contacts,
  };
}

export async function employeeRoutes(app: FastifyInstance) {
  app.get("/api/employees", async (request, reply) => {
    const parsed = listQuerySchema.safeParse(request.query);

    if (!parsed.success) {
      return reply.code(400).send({
        error: "invalid_query",
        details: parsed.error.flatten(),
      });
    }

    const { query, department, status, limit, offset } = parsed.data;
    const searchPattern = query ? `%${query}%` : null;

    const sql = `
      ${employeeSelectSql}
      WHERE (
        $1::text IS NULL
        OR employee.full_name ILIKE $1
        OR employee.employee_code ILIKE $1
        OR position.title ILIKE $1
        OR EXISTS (
          SELECT 1
          FROM employee_contacts AS searchable_contact
          WHERE searchable_contact.employee_id = employee.id
            AND searchable_contact.is_active = TRUE
            AND searchable_contact.contact_value ILIKE $1
        )
      )
      AND (
        $2::text IS NULL
        OR department.code = $2
        OR department.name = $2
      )
      AND (
        $3::text IS NULL
        OR employee.status = $3
      )
      ORDER BY employee.full_name
      LIMIT $4
      OFFSET $5
    `;

    const result = await pool.query<EmployeeRow>(sql, [
      searchPattern,
      department ?? null,
      status ?? null,
      limit,
      offset,
    ]);

    return {
      items: result.rows.map(mapEmployee),
      pagination: {
        total: result.rows[0]?.total_count ?? 0,
        limit,
        offset,
      },
    };
  });

  app.get("/api/employees/:id", async (request, reply) => {
    const parsed = employeeParamsSchema.safeParse(request.params);

    if (!parsed.success) {
      return reply.code(400).send({
        error: "invalid_employee_id",
      });
    }

    const sql = `
      ${employeeSelectSql}
      WHERE employee.id::text = $1
         OR employee.employee_code = $1
      LIMIT 1
    `;
    const result = await pool.query<EmployeeRow>(sql, [parsed.data.id]);
    const employee = result.rows[0];

    if (!employee) {
      return reply.code(404).send({
        error: "employee_not_found",
      });
    }

    return mapEmployee(employee);
  });
}
