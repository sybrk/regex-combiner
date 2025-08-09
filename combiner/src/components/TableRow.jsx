import { flexRender } from "@tanstack/react-table";
import { memo } from "react";

const TableRow = memo(({ row }) => {
    console.log("myrow", row); // This will now only log when this specific row actually changes
    
    return (
      <tr key={row.id} className="hover:bg-gray-50">
        {row.getVisibleCells().map(cell => (
          <td key={cell.id} className="px-6 py-4 whitespace-nowrap">
            {flexRender(
              cell.column.columnDef.cell,
              cell.getContext()
            )}
          </td>
        ))}
      </tr>
    );
  });

export default TableRow