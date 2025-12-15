import nrsLogo from "@/assets/logos/NRS-logo.png";
import Image from "next/image";

export function Logo() {
  return (
    <div className="relative h-8 max-w-[10.847rem]">
      <Image
        src={nrsLogo}
        alt="Nigeria Revenue Service logo"
        role="presentation"
        quality={100}
        fill
        className="object-contain"
      />
    </div>
  );
}
