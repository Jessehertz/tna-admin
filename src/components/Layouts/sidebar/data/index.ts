import * as Icons from "../icons";

export const NAV_DATA = [
  {
    label: "MAIN MENU",
    items: [
      {
        title: "Dashboard",
        icon: Icons.HomeIcon,
        url: "/",

        items: [],
      },

      {
        title: "Edit Profile",
        icon: Icons.User,
        url: "/edit-profile",

        items: [],
      },
    ],
  },
  {
    label: "MODULE MANAGEMENT",
    items: [
      {
        title: "Upload New Module",
        icon: Icons.HomeIcon,
        url: "/upload-new-module",

        items: [],
      },

      {
        title: "Manage Modules",
        icon: Icons.User,
        url: "/manage-modules",

        items: [],
      },
    ],
  },
];
