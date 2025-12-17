import { PaymentsOverview } from "@/components/Charts/payments-overview";
import { UsedDevices } from "@/components/Charts/used-devices";
import { WeeksProfit } from "@/components/Charts/weeks-profit";
import { TopChannels } from "@/components/Tables/top-channels";
import { TopChannelsSkeleton } from "@/components/Tables/top-channels/skeleton";
import { Button } from "@/components/ui-elements/button";
import { createTimeFrameExtractor } from "@/utils/timeframe-extractor";
import { Suspense } from "react";
import { ChatsCard } from "./_components/chats-card";
import { OverviewCardsGroup } from "./_components/overview-cards";
import { OverviewCardsSkeleton } from "./_components/overview-cards/skeleton";
import { RegionLabels } from "./_components/region-labels";

type PropsType = {
  searchParams: Promise<{
    selected_time_frame?: string;
  }>;
};

export default async function Home({ searchParams }: PropsType) {
  const { selected_time_frame } = await searchParams;
  const extractTimeFrame = createTimeFrameExtractor(selected_time_frame);

  return (
    <>
      <Suspense fallback={<OverviewCardsSkeleton />}>
        <OverviewCardsGroup />
      </Suspense>

      <div className="mt-4 flex flex-wrap justify-between gap-4 md:mt-6 md:gap-6">
        <Button
          label="Add / Manage SME's"
          variant="primary"
          shape="rounded"
          className="min-w-[200px] flex-1 !bg-[rgb(212,53,34)] hover:!bg-[rgb(212,53,34)]/90"
        />
        <Button
          label="Set Training Cycle"
          variant="primary"
          shape="rounded"
          className="min-w-[200px] flex-1 !bg-[rgb(212,53,34)] hover:!bg-[rgb(212,53,34)]/90"
        />
        <Button
          label="View Reports"
          variant="primary"
          shape="rounded"
          className="min-w-[200px] flex-1 !bg-[rgb(212,53,34)] hover:!bg-[rgb(212,53,34)]/90"
        />
      </div>
      <div className="mt-4 gap-4 md:mt-6 md:gap-6 2xl:mt-9 2xl:gap-7.5">
        <div className="col-span-12 grid xl:col-span-8">
          <Suspense fallback={<TopChannelsSkeleton />}>
            <TopChannels />
          </Suspense>
        </div>
      </div>
    </>
  );
}
