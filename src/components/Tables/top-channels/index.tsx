import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { compactFormat, standardFormat } from "@/lib/format-number";
import { cn } from "@/lib/utils";
import Image from "next/image";
import Link from "next/link";
import { getTopChannels } from "../fetch";
import dayjs from "dayjs";
import { TrashIcon } from "@/assets/icons";
import { DownloadIcon, PreviewIcon } from "../icons";

export async function TopChannels({ className }: { className?: string }) {
  const data = await getTopChannels();

  return (
    <div
      className={cn(
        "grid rounded-[10px] bg-white px-7.5 pb-4 pt-7.5 shadow-1 dark:bg-gray-dark dark:shadow-card",
        className,
      )}
    >
      <h2 className="mb-4 text-body-2xlg font-bold text-dark dark:text-white">
        Featured Modules
      </h2>

      <Table>
        <TableHeader>
          <TableRow className="border-none uppercase [&>th]:text-center">
            <TableHead className="w-16 text-center">S/N</TableHead>
            <TableHead>Module Name</TableHead>
            <TableHead>Status</TableHead>
            <TableHead>Date Created</TableHead>
            <TableHead>Action</TableHead>
          </TableRow>
        </TableHeader>

        <TableBody>
          {data.map((channel, i) => (
            <TableRow
              className="text-center text-base font-medium text-dark dark:text-white"
              key={channel.name + i}
            >
              <TableCell className="w-16 text-center">{i + 1}</TableCell>

              <TableCell>{channel.name}</TableCell>

              <TableCell
                className={cn("text-center", {
                  "text-green-light-1": channel.status === "Active",
                  "text-red-light-1": channel.status === "Inactive",
                })}
              >
                <div
                  className={cn(
                    "rounded-full px-3.5 py-1 text-sm font-medium",
                    {
                      "bg-[#219653]/[0.08] text-[#219653]":
                        channel.status === "Active",
                      "bg-[#D34053]/[0.08] text-[#D34053]":
                        channel.status === "Inactive",
                    },
                  )}
                >
                  {channel.status}
                </div>
              </TableCell>
              <TableCell className="text-center">
                {dayjs(channel.dateCreated).format("MMM DD, YYYY")}
              </TableCell>

              <TableCell className="text-center">
                <div className="flex items-center justify-center gap-x-3.5">
                  <Link
                    href={`/manage-modules-view?id=${channel.name || i + 1}`}
                    className="hover:text-primary"
                  >
                    <span className="sr-only">View Module</span>
                    <PreviewIcon />
                  </Link>

                  <button className="hover:text-primary">
                    <span className="sr-only">Delete Module</span>
                    <TrashIcon />
                  </button>
                </div>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
}
