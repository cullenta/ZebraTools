import React, { useEffect, useState } from 'react';
import axios from 'axios';

export default function AttendancePage() {
  const [students, setStudents] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');

  useEffect(() => {
    axios.get('http://localhost:8000/students/')
      .then(response => {
        setStudents(response.data);
      })
      .catch(error => {
        console.error('Error fetching data:', error);
      });
  }, []);

  const filteredStudents = students.filter(student =>
    student.name.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="p-6 max-w-4xl mx-auto">
      <h1 className="text-red-500 text-3xl font-bold mb-4">Students</h1>
      <input
        type="text"
        placeholder="Search by name..."
        value={searchTerm}
        onChange={e => setSearchTerm(e.target.value)}
        className="w-full p-2 border border-gray-300 rounded mb-6"
      />

      <div className="grid grid-cols-1 gap-4">
        {filteredStudents.map((student, index) => (

          <div key={index} className="p-4 flex flex-col">
            <p><strong>Name:</strong> {student.name}</p>
            <p><strong>Username:</strong> {student.lmsusername}@zebrarobotics.com</p>
            <p><strong>Password:</strong> {student.lmspassword}</p>
            <p><strong>Course:</strong> {student.course}</p>
          </div>
        ))}
      </div>
    </div>
  );
}
