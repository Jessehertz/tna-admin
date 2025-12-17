"use client";

import { createContext, useContext, useState, useEffect } from "react";

type UserRole = "sme" | "la";

type UserRoleContextType = {
  user_role: UserRole;
  setUserRole: (role: UserRole) => void;
};

const UserRoleContext = createContext<UserRoleContextType | undefined>(
  undefined,
);

export function useUserRole() {
  const context = useContext(UserRoleContext);
  if (context === undefined) {
    throw new Error("useUserRole must be used within a UserRoleProvider");
  }
  return context;
}

export function UserRoleProvider({
  children,
  defaultRole = "la",
}: {
  children: React.ReactNode;
  defaultRole?: UserRole;
}) {
  // Initialize state from localStorage or use defaultRole
  const [user_role, setUser_role] = useState<UserRole>(() => {
    // Only access localStorage on client side
    if (typeof window !== "undefined") {
      const storedRole = localStorage.getItem("user_role") as UserRole | null;
      if (storedRole && (storedRole === "sme" || storedRole === "la")) {
        return storedRole;
      }
    }
    return defaultRole;
  });

  // Save user role to localStorage whenever it changes
  useEffect(() => {
    if (typeof window !== "undefined") {
      localStorage.setItem("user_role", user_role);
      console.log("User role updated to:", user_role);
    }
  }, [user_role]);

  const setUserRole = (role: UserRole) => {
    setUser_role(role);
  };

  return (
    <UserRoleContext.Provider
      value={{
        user_role,
        setUserRole,
      }}
    >
      {children}
    </UserRoleContext.Provider>
  );
}
