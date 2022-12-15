// import React from 'react';
// import { render, screen } from '@testing-library/react';
// import App from './App';

// test('renders learn react link', () => {
//   // render(<App />);
//   const linkElement = screen.getByText(/learn react/i);
//   expect(linkElement).toBeInTheDocument();
// });

import moment from "moment";

test('test months',()=>{
  const start = moment('201910','YYYYMM');
  const end = moment('202210','YYYYMM');

  const months = [start];
  while (months[months.length - 1].diff(end) < 0) {
    months.push(months[months.length - 1].clone().add(1, 'month'));
  }

  console.log(months.map(m=>m.format('YYYYMM')));
})
