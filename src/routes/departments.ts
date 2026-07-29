import type { FastifyInstance } from "fastify";
import { pool } from "../lib/db.js";

type DepartmentRow = {
  id: string;
  code: string;
  name: string;
  location: string | null;
  position_count: number;
  employee_count: number;
  vacancy_count: number;
};

const departmentsSql = `
  SELECT
    department.id::text,
    department.code,
    department.name,
    department.location,
    count(position.id)::int AS position_count,
    count(assignment.id)::int AS employee_count,
    count(position.id) FILTER (
      WHERE assignment.id IS NULL
    )::int AS vacancy_count
  FROM departments AS department
  LEFT JOIN positions AS position
    ON position.department_id = department.id
   AND position.is_active = TRUE
  LEFT JOIN employment_assignments AS assignment
    ON assignment.position_id = position.id
   AND assignment.ended_on IS NULL
  WHERE department.is_active = TRUE
  GROUP BY
    department.id,
    department.code,
    department.name,
    department.location
  ORDER BY department.name
`;

export async function departmentRoutes(app: FastifyInstance) {
  app.get("/api/departments", async () => {
    const result = await pool.query<DepartmentRow>(departmentsSql);

    return {
      items: result.rows.map((row) => ({
        id: row.id,
        code: row.code,
        name: row.name,
        location: row.location,
        positionCount: row.position_count,
        employeeCount: row.employee_count,
        vacancyCount: row.vacancy_count,
      })),
      total: result.rowCount ?? 0,
    };
  });
}
