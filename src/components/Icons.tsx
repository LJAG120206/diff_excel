import type { SVGProps } from 'react'

type IconProps = SVGProps<SVGSVGElement>

export function IconTable(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1.5" stroke="currentColor" />
      <path d="M1.5 6.5h13M1.5 10.5h13M6 2.5v11" stroke="currentColor" />
    </svg>
  )
}

export function IconPlus(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M8 3v10M3 8h10" stroke="currentColor" strokeWidth="1.5" />
    </svg>
  )
}

export function IconMinus(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M3 8h10" stroke="currentColor" strokeWidth="1.5" />
    </svg>
  )
}

export function IconSwap(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path
        d="M3 6h10M10 3l3 3-3 3M13 10H3M6 13l-3-3 3-3"
        stroke="currentColor"
        strokeWidth="1.2"
        strokeLinejoin="round"
      />
    </svg>
  )
}

export function IconSun(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <circle cx="8" cy="8" r="2.5" stroke="currentColor" />
      <path
        d="M8 1.5v1.5M8 13v1.5M1.5 8H3M13 8h1.5M3.4 3.4l1.1 1.1M11.5 11.5l1.1 1.1M3.4 12.6l1.1-1.1M11.5 4.5l1.1-1.1"
        stroke="currentColor"
        strokeLinecap="round"
      />
    </svg>
  )
}

export function IconMoon(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path
        d="M13 10.5A5.5 5.5 0 1 1 5.5 3 4.5 4.5 0 0 0 13 10.5z"
        stroke="currentColor"
        strokeLinejoin="round"
      />
    </svg>
  )
}

export function IconGear(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <circle cx="8" cy="8" r="2" stroke="currentColor" />
      <path
        d="M8 2.5v1.2M8 12.3V13.5M2.5 8h1.2M12.3 8h1.2M4.1 4.1l.9.9M11 11l.9.9M4.1 11.9l.9-.9M11 5l.9-.9"
        stroke="currentColor"
        strokeLinecap="round"
      />
    </svg>
  )
}

export function IconDownload(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M8 3v7M5 8l3 3 3-3M3 12.5h10" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  )
}

export function IconArrowUp(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M8 12V4M4.5 7.5 8 4l3.5 3.5" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}

export function IconArrowDown(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M8 4v8M4.5 8.5 8 12l3.5-3.5" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}

export function IconFile(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M4 2.5h5.5L12.5 6v7.5H4v-11z" stroke="currentColor" strokeLinejoin="round" />
      <path d="M9.5 2.5V6h3" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}

export function IconReset(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M4 6.5A5 5 0 1 1 3.5 10" stroke="currentColor" />
      <path d="M4 3.5v3h3" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}

export function IconMerge(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M5 2v7.5A2.5 2.5 0 0 0 7.5 12H11" stroke="currentColor" />
      <circle cx="5" cy="2" r="1.3" fill="currentColor" />
      <circle cx="5" cy="14" r="1.3" fill="currentColor" />
      <path d="M11 9.5 13.5 12 11 14.5" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}

export function IconChevronLeft(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M10 3.5 5.5 8 10 12.5" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}

export function IconChevronRight(props: IconProps) {
  return (
    <svg viewBox="0 0 16 16" fill="none" aria-hidden="true" {...props}>
      <path d="M6 3.5 10.5 8 6 12.5" stroke="currentColor" strokeLinejoin="round" />
    </svg>
  )
}
