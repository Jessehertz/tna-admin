"use client";

import { SidebarProvider } from "@/components/Layouts/sidebar/sidebar-context";
import { UserRoleProvider } from "@/contexts/user-role-context";
import { ThemeProvider } from "next-themes";

export function Providers({ children }: { children: React.ReactNode }) {
  return (
    <ThemeProvider defaultTheme="light" attribute="class">
      <UserRoleProvider defaultRole="sme">
        <SidebarProvider>{children}</SidebarProvider>
      </UserRoleProvider>
    </ThemeProvider>
  );
}
