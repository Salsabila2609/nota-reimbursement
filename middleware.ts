import { NextResponse } from 'next/server'
import type { NextRequest } from 'next/server'

export function middleware(request: NextRequest) {
  const token = request.cookies.get('auth-token')
  const { pathname } = request.nextUrl

  // Public routes
  if (pathname === '/' || pathname.startsWith('/api/auth')) {
    return NextResponse.next()
  }

  // Protect /driver and /admin pages
  if ((pathname.startsWith('/driver') || pathname.startsWith('/admin')) && !token) {
    return NextResponse.redirect(new URL('/', request.url))
  }

  return NextResponse.next()
}

export const config = {
  matcher: ['/driver/:path*', '/admin/:path*', '/api/submissions/:path*', '/api/generate-pdf/:path*', '/api/drivers/:path*']
}
