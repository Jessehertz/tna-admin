"use client";

import { Header } from "@/components/Layouts/header";
import { Sidebar } from "@/components/Layouts/sidebar";
import { usePathname } from "next/navigation";
import type { PropsWithChildren } from "react";

export function ConditionalLayout({ children }: PropsWithChildren) {
  const pathname = usePathname();
  const isAuthPage = pathname?.startsWith("/auth");

  if (isAuthPage) {
    return (
      <div className="min-h-screen bg-gray-2 dark:bg-[#020d1a]">
        <main className="flex min-h-screen items-center justify-center p-4 md:p-6 2xl:p-10">
          {children}
        </main>
      </div>
    );
  }

  return (
    <div className="flex min-h-screen">
      <Sidebar />

      <div className="flex w-full flex-col bg-gray-2 dark:bg-[#020d1a]">
        <Header />

        <main className="isolate mx-auto w-full max-w-screen-2xl flex-1 overflow-hidden p-4 md:p-6 2xl:p-10">
          {children}
        </main>

        <footer className="border-t border-stroke bg-white px-4 py-5 dark:border-stroke-dark dark:bg-gray-dark md:px-5 2xl:px-10">
          <p className="text-center text-sm font-medium text-dark-4 dark:text-dark-6">
            © 2025 Federal Inland Revenue Service. All Rights Reserved.
          </p>
        </footer>
      </div>
    </div>
  );
}
