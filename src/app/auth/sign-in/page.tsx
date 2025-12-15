import Signin from "@/components/Auth/Signin";
import type { Metadata } from "next";
import Image from "next/image";
import Link from "next/link";
import NRSLogo from "@/assets/logos/NRS-logo.png";

export const metadata: Metadata = {
  title: "Sign in",
};

export default function SignIn() {
  return (
    <div className="rounded-[10px] bg-white shadow-1 dark:bg-gray-dark dark:shadow-card">
      <div className="flex flex-wrap items-center">
        <div className="w-full xl:w-1/2">
          <div className="w-full p-4 sm:p-12.5 xl:p-15">
            <Signin />
          </div>
        </div>

        <div className="hidden w-full p-7.5 xl:block xl:w-1/2">
          <div className="overflow-hidden rounded-2xl bg-gradient-to-br from-black via-gray-800 to-gray-700 px-12.5 pt-12.5">
            <Link className="mb-10 inline-block" href="/">
              <Image src={NRSLogo} alt="NRS Logo" width={176} height={32} />
            </Link>
            <p className="mb-3 text-xl font-medium text-white">
              Sign in to your account
            </p>

            <h1 className="mb-4 text-2xl font-bold text-white sm:text-heading-3">
              Welcome Back!
            </h1>

            <p className="w-full max-w-[375px] font-medium text-white">
              Please sign in to your account with your active directory
              credentials
            </p>

            <div className="mt-31">
              <Image
                src={"/images/grids/grid-02.svg"}
                alt="Logo"
                width={405}
                height={325}
                className="mx-auto dark:opacity-30"
              />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
