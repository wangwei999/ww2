import type { Metadata } from 'next';
import { Inspector } from 'react-dev-inspector';
import './globals.css';

export const metadata: Metadata = {
  title: {
    default: '授信数据填充 | 扣子编程',
    template: '%s | 扣子编程',
  },
  description:
    '授信数据填充工具，支持PDF识别、智能匹配、自动填充授信数据到Excel表格。',
  keywords: [
    '授信数据',
    '数据填充',
    'PDF识别',
    'Excel处理',
    '扣子编程',
    'Coze Code',
  ],
  authors: [{ name: 'Coze Code Team', url: 'https://code.coze.cn' }],
  generator: 'Coze Code',
  // icons: {
  //   icon: '',
  // },
  openGraph: {
    title: '授信数据填充 | 扣子编程',
    description:
      '授信数据填充工具，支持PDF识别、智能匹配、自动填充授信数据到Excel表格。',
    url: 'https://code.coze.cn',
    siteName: '授信数据填充',
    locale: 'zh_CN',
    type: 'website',
    // images: [
    //   {
    //     url: '',
    //     width: 1200,
    //     height: 630,
    //     alt: '扣子编程 - 你的 AI 工程师',
    //   },
    // ],
  },
  // twitter: {
  //   card: 'summary_large_image',
  //   title: 'Coze Code | Your AI Engineer is Here',
  //   description:
  //     'Build and deploy full-stack applications through AI conversation. No env setup, just flow.',
  //   // images: [''],
  // },
  robots: {
    index: true,
    follow: true,
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  const isDev = process.env.NODE_ENV === 'development';

  return (
    <html lang="en">
      <body className={`antialiased`}>
        {isDev && <Inspector />}
        {children}
      </body>
    </html>
  );
}
