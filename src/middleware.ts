import { NextResponse } from 'next/server'
import type { NextRequest } from 'next/server'

// NOTE: This is a placeholder middleware.
// Full backend authentication, RBAC, and session management are out of scope for the current phase.

export function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;

  // Log access for observation, but no redirection logic will be applied
  // as authentication is not implemented.
  console.log(`Middleware: Accessing path ${pathname}.`);

  // Allow all requests to proceed
  return NextResponse.next();
}

// Configure which paths the middleware should run on
export const config = {
  matcher: [
    /*
     * Match all request paths except for the ones starting with:
     * - api (API routes)
     * - _next/static (static files)
     * - _next/image (image optimization files)
     * - favicon.ico (favicon file)
     */
    '/((?!api|_next/static|_next/image|favicon.ico).*)',
  ],
}
