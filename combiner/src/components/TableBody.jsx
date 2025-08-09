import { memo } from "react";
import TableRow from "./TableRow";

const TableBody = memo(({ rows }) => {
    return (
      <tbody className="bg-white divide-y divide-gray-200">
        {rows.map(row => (
          <TableRow key={row.id} row={row} />
        ))}
      </tbody>
    );
  });

  export default TableBody