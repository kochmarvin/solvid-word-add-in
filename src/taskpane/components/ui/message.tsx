import * as React from "react";

import { cn } from "@/lib/utils";

function Message({
  className,
  children,
  role = "user",
  ...props
}: React.ComponentProps<"div"> & { role?: "ai" | "user" }) {
  return (
    <>
      {role == "ai" ? (
        <div
          className={cn(
            "max-w-[60%] md:max-w-[60%] max-md:max-w-[85%] rounded-lg px-2.5 py-2 text-base md:text-sm flex field-sizing-content",
            className
          )}
          {...props}
        >
          {children}
        </div>
      ) : <div
        className={cn(
          "self-end max-w-[60%] md:max-w-[60%] max-md:max-w-[85%] rounded-lg bg-[#e8f5f3] px-2.5 py-2 text-base md:text-sm flex field-sizing-content",
          className
        )}
        {...props}
      >
        {children}
      </div>
      }

    </>
  );
}

function Conversation({ className, children, ...props }: React.ComponentProps<"div">) {
  return (
    <div className={cn("flex flex-col w-full gap-2", className)} {...props}>
      {children}
    </div>
  );
}

export { Message, Conversation };
