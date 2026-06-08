import { Check } from "lucide-react";
import clsx from "clsx";

interface Props {
  steps: string[];
  current: number; // 1-based
  onJump?: (step: number) => void;
  maxReached: number;
}

export function Stepper({ steps, current, onJump, maxReached }: Props) {
  return (
    <div className="flex items-center justify-center gap-2 sm:gap-4">
      {steps.map((label, i) => {
        const n = i + 1;
        const done = n < current;
        const active = n === current;
        const reachable = n <= maxReached;
        return (
          <div key={label} className="flex items-center gap-2 sm:gap-4">
            <button
              disabled={!reachable}
              onClick={() => reachable && onJump?.(n)}
              className={clsx(
                "flex items-center gap-2 rounded-full px-3 py-1.5 text-sm font-medium transition",
                active && "bg-brand-600 text-white shadow-lg shadow-brand-600/30",
                done && "text-brand-300 hover:bg-white/5",
                !active && !done && "text-slate-500",
                reachable && !active && "cursor-pointer hover:text-slate-300",
                !reachable && "cursor-not-allowed"
              )}
            >
              <span
                className={clsx(
                  "flex h-6 w-6 items-center justify-center rounded-full border text-xs",
                  active && "border-white bg-white/20",
                  done && "border-brand-400 bg-brand-500/20",
                  !active && !done && "border-slate-700"
                )}
              >
                {done ? <Check size={14} /> : n}
              </span>
              <span className="hidden sm:inline">{label}</span>
            </button>
            {n < steps.length && (
              <div
                className={clsx(
                  "h-px w-6 sm:w-12 transition",
                  n < current ? "bg-brand-500" : "bg-slate-700"
                )}
              />
            )}
          </div>
        );
      })}
    </div>
  );
}
