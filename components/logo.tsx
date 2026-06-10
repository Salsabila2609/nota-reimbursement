// src/components/Logo.tsx
// Taruh file logo "Logo IOH-Color.png" di: public/logo-ioh.png
// ATAU di src/components/Logo IOH-Color.png dan import langsung

import Image from 'next/image'

type LogoProps = {
  height?: number
  className?: string
}

export default function Logo({ height = 32, className }: LogoProps) {
  return (
    <Image
      src="/logo-ioh.png"          // letakkan Logo IOH-Color.png di folder /public/ dengan nama logo-ioh.png
      alt="Indosat Ooredoo Hutchison"
      height={height}
      width={height * 3.2}         // aspect ratio logo IOH ± 3.2 : 1
      style={{ objectFit: 'contain', display: 'block' }}
      className={className}
      priority
    />
  )
}