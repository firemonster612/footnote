import { ChevronDown, type LucideIcon } from "lucide-react";
import type { ButtonHTMLAttributes, InputHTMLAttributes, SelectHTMLAttributes } from "react";

type ButtonVariant = "primary" | "secondary" | "danger";

const buttonVariants: Record<ButtonVariant, string> = {
  primary: "bg-accent text-white hover:bg-accent-hover",
  secondary:
    "border border-neutral-300 bg-white text-neutral-800 hover:bg-neutral-50 dark:border-neutral-600 dark:bg-neutral-800 dark:text-neutral-100 dark:hover:bg-neutral-700",
  danger:
    "border border-neutral-300 bg-white text-red-700 hover:bg-red-50 dark:border-neutral-600 dark:bg-neutral-800 dark:text-red-400 dark:hover:bg-red-950",
};

export function Button({
  variant = "secondary",
  className = "",
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & { variant?: ButtonVariant }) {
  return (
    <button
      type="button"
      className={`inline-flex h-7 items-center justify-center gap-1.5 rounded-md px-2.5 text-[12px] font-medium whitespace-nowrap disabled:pointer-events-none disabled:opacity-50 ${buttonVariants[variant]} ${className}`}
      {...props}
    />
  );
}

export function IconButton({
  icon: Icon,
  label,
  className = "",
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & { icon: LucideIcon; label: string }) {
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      className={`inline-flex size-7 shrink-0 items-center justify-center rounded-md text-neutral-600 hover:bg-neutral-100 hover:text-neutral-900 disabled:pointer-events-none disabled:opacity-40 dark:text-neutral-400 dark:hover:bg-neutral-800 dark:hover:text-neutral-100 ${className}`}
      {...props}
    >
      <Icon size={16} />
    </button>
  );
}

const selectVariants = {
  ghost:
    "h-7 bg-transparent text-[12px] text-neutral-700 hover:bg-neutral-100 dark:text-neutral-300 dark:hover:bg-neutral-800",
  field:
    "h-8 border border-neutral-300 bg-white text-[13px] dark:border-neutral-600 dark:bg-neutral-800",
};

export function Select({
  variant = "ghost",
  className = "",
  children,
  ...props
}: SelectHTMLAttributes<HTMLSelectElement> & { variant?: keyof typeof selectVariants }) {
  return (
    <span className={`relative inline-flex min-w-0 items-center ${className}`}>
      <select
        className={`w-full min-w-0 appearance-none truncate rounded-md pr-6 pl-2 disabled:opacity-50 ${selectVariants[variant]}`}
        {...props}
      >
        {children}
      </select>
      <ChevronDown size={12} className="pointer-events-none absolute right-1.5 text-neutral-500" />
    </span>
  );
}

export function TextInput({ className = "", ...props }: InputHTMLAttributes<HTMLInputElement>) {
  return (
    <input
      className={`h-8 w-full min-w-0 rounded-md border border-neutral-300 bg-white px-2 text-[13px] placeholder:text-neutral-400 dark:border-neutral-600 dark:bg-neutral-800 ${className}`}
      {...props}
    />
  );
}
